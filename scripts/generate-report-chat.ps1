[CmdletBinding()]
param(
  [string]$PromptPath,

  [string]$PromptText,

  [string[]]$ReferenceTextPaths,

  [string[]]$ReferenceUrls,

  [string]$OutFile,

  [string]$SessionKey = "agent:gpt:main",

  [string]$BrowserProfile = $env:OPENCLAW_BROWSER_PROFILE,

  [string]$OpenClawCmd = $env:OPENCLAW_CMD,

  [int]$ReferenceMaxChars = 30000,

  [switch]$SkipSessionReset
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-ConfigValue {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Object,

    [Parameter(Mandatory = $true)]
    [string]$Name
  )

  if ($null -eq $Object) {
    return $null
  }

  if ($Object.PSObject.Properties.Name -contains $Name) {
    return $Object.$Name
  }

  return $null
}

if ([string]::IsNullOrWhiteSpace($PromptPath) -eq [string]::IsNullOrWhiteSpace($PromptText)) {
  throw "Provide exactly one of -PromptPath or -PromptText."
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$configPath = Join-Path (Join-Path $HOME ".openclaw") "openclaw.json"
$resolvedConfigPath = (Resolve-Path -LiteralPath $configPath).Path
$config = (Get-Content -LiteralPath $resolvedConfigPath -Raw -Encoding UTF8) | ConvertFrom-Json

$gatewayHost = "127.0.0.1"
$gatewayPort = 18789
$gateway = Get-ConfigValue -Object $config -Name "gateway"
if ($null -ne (Get-ConfigValue -Object $gateway -Name "bindHost")) {
  $gatewayHost = [string](Get-ConfigValue -Object $gateway -Name "bindHost")
}
if ($null -ne (Get-ConfigValue -Object $gateway -Name "port") -and -not [string]::IsNullOrWhiteSpace([string](Get-ConfigValue -Object $gateway -Name "port"))) {
  $gatewayPort = [int](Get-ConfigValue -Object $gateway -Name "port")
}
$gatewayAuth = Get-ConfigValue -Object $gateway -Name "auth"
$gatewayToken = [string](Get-ConfigValue -Object $gatewayAuth -Name "token")
if ([string]::IsNullOrWhiteSpace($gatewayToken)) {
  throw "OpenClaw gateway token is missing in $resolvedConfigPath"
}

$nodeCommand = Get-Command node -ErrorAction SilentlyContinue
if ($null -eq $nodeCommand -or [string]::IsNullOrWhiteSpace($nodeCommand.Source)) {
  throw "Node.js is required to send chat requests through the OpenClaw gateway."
}

$userPrompt = if (-not [string]::IsNullOrWhiteSpace($PromptPath)) {
  Get-Content -LiteralPath (Resolve-Path -LiteralPath $PromptPath).Path -Raw -Encoding UTF8
} else {
  $PromptText
}

if ([string]::IsNullOrWhiteSpace($OutFile)) {
  $OutFile = Join-Path $repoRoot ("tests-output\generated-report-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".txt")
}

$resolvedOutFile = [System.IO.Path]::GetFullPath($OutFile)
$outDir = Split-Path -Parent $resolvedOutFile
if (-not [string]::IsNullOrWhiteSpace($outDir)) {
  New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

$preparedPromptPath = Join-Path $outDir (([System.IO.Path]::GetFileNameWithoutExtension($resolvedOutFile)) + ".prepared-prompt.txt")
$preparedPromptResult = & (Join-Path $repoRoot "scripts\prepare-report-prompt.ps1") `
  -PromptText $userPrompt `
  -OutFile $preparedPromptPath `
  -ReferenceTextPaths $ReferenceTextPaths `
  -ReferenceUrls $ReferenceUrls `
  -BrowserProfile $BrowserProfile `
  -OpenClawCmd $OpenClawCmd `
  -ReferenceMaxChars $ReferenceMaxChars
$userPrompt = Get-Content -LiteralPath $preparedPromptPath -Raw -Encoding UTF8

$skillMarkdown = Get-Content -LiteralPath (Join-Path $repoRoot "SKILL.md") -Raw -Encoding UTF8
$skillBody = [regex]::Replace($skillMarkdown, '^(?s)---\s*.*?\s*---\s*', '')
$messageText = @(
  "Use the following task policy as authoritative instructions for this turn."
  ""
  $skillBody.Trim()
  ""
  "User request:"
  $userPrompt.Trim()
  ""
  "Return only the final Chinese experiment report body. Prefer a complete, moderately detailed, submission-ready report instead of a terse outline. Do not ask follow-up questions. Do not output explanations or meta commentary."
) -join [Environment]::NewLine

$messageFile = Join-Path $outDir (([System.IO.Path]::GetFileNameWithoutExtension($resolvedOutFile)) + ".prompt.txt")
[System.IO.File]::WriteAllText($messageFile, $messageText, (New-Object System.Text.UTF8Encoding($true)))

if (-not $SkipSessionReset) {
  & (Join-Path $repoRoot "scripts\reset-openclaw-session.ps1") -SessionKey $SessionKey | Out-Null
}

$nodeScript = @'
const crypto = require("crypto");
const fs = require("fs");
const WS = globalThis.WebSocket || require("ws");

const token = process.env.OC_CHAT_GATEWAY_TOKEN;
const gatewayUrl = process.env.OC_CHAT_GATEWAY_URL;
const sessionKey = process.env.OC_CHAT_SESSION_KEY;
const messagePath = process.env.OC_CHAT_MESSAGE_PATH;
const outPath = process.env.OC_CHAT_OUT_PATH;

if (!token || !gatewayUrl || !sessionKey || !messagePath || !outPath) {
  throw new Error("Missing gateway chat environment.");
}

const messageText = fs.readFileSync(messagePath, "utf8");
const ws = new WS(gatewayUrl);
const pending = new Map();
let finalWritten = false;

function request(method, params) {
  return new Promise((resolve, reject) => {
    const id = crypto.randomUUID();
    pending.set(id, { resolve, reject });
    ws.send(JSON.stringify({ type: "req", id, method, params }));
  });
}

function extractText(message) {
  if (!message || typeof message !== "object") return "";
  if (typeof message.text === "string") return message.text;
  if (Array.isArray(message.content)) {
    return message.content
      .map((part) => (part && typeof part.text === "string") ? part.text : "")
      .filter(Boolean)
      .join("\n");
  }
  return "";
}

function rejectPending(message) {
  for (const slot of pending.values()) {
    slot.reject(new Error(message));
  }
  pending.clear();
}

ws.onopen = async () => {
  try {
    await request("connect", {
      minProtocol: 3,
      maxProtocol: 3,
      client: {
        id: "cli",
        version: "codex-gateway-chat",
        platform: process.platform,
        mode: "cli",
        instanceId: crypto.randomUUID()
      },
      role: "operator",
      scopes: ["operator.admin", "operator.approvals", "operator.pairing"],
      auth: { token },
      userAgent: "codex-gateway-chat",
      locale: "zh-CN"
    });
    await request("chat.send", {
      sessionKey,
      message: messageText,
      deliver: false,
      idempotencyKey: crypto.randomUUID()
    });
  } catch (error) {
    process.stderr.write((error && error.stack) ? error.stack : String(error));
    process.exitCode = 1;
    ws.close();
  }
};

ws.onmessage = (event) => {
  const raw = typeof event.data === "string" ? event.data : event.data.toString();
  const message = JSON.parse(raw);

  if (message.type === "res") {
    const slot = pending.get(message.id);
    if (!slot) return;
    pending.delete(message.id);
    if (message.ok) slot.resolve(message.payload);
    else slot.reject(new Error((message.error && message.error.message) || "request failed"));
    return;
  }

  if (message.type === "event" && message.event === "chat" && message.payload && message.payload.state === "final") {
    const text = extractText(message.payload.message).trim();
    fs.writeFileSync(outPath, text, "utf8");
    process.stdout.write(text);
    finalWritten = true;
    ws.close();
  }
};

ws.onerror = (error) => {
  process.stderr.write((error && error.message) ? error.message : String(error));
};

ws.onclose = () => {
  if (!finalWritten && process.exitCode !== 1) {
    rejectPending("socket closed before final response");
    process.exitCode = 1;
  } else {
    pending.clear();
  }
};
'@

$previousToken = $env:OC_CHAT_GATEWAY_TOKEN
$previousUrl = $env:OC_CHAT_GATEWAY_URL
$previousKey = $env:OC_CHAT_SESSION_KEY
$previousMessagePath = $env:OC_CHAT_MESSAGE_PATH
$previousOutPath = $env:OC_CHAT_OUT_PATH

try {
  $env:OC_CHAT_GATEWAY_TOKEN = $gatewayToken
  $env:OC_CHAT_GATEWAY_URL = "ws://{0}:{1}" -f $gatewayHost, $gatewayPort
  $env:OC_CHAT_SESSION_KEY = $SessionKey
  $env:OC_CHAT_MESSAGE_PATH = $messageFile
  $env:OC_CHAT_OUT_PATH = $resolvedOutFile
  $output = $nodeScript | & $nodeCommand.Source -
  if ($LASTEXITCODE -ne 0) {
    throw "OpenClaw gateway chat generation failed."
  }
  Write-Output $output
} finally {
  $env:OC_CHAT_GATEWAY_TOKEN = $previousToken
  $env:OC_CHAT_GATEWAY_URL = $previousUrl
  $env:OC_CHAT_SESSION_KEY = $previousKey
  $env:OC_CHAT_MESSAGE_PATH = $previousMessagePath
  $env:OC_CHAT_OUT_PATH = $previousOutPath
}
