[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$SessionKey,

  [string]$ConfigPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
  $ConfigPath = Join-Path (Join-Path $HOME ".openclaw") "openclaw.json"
}

$resolvedConfigPath = (Resolve-Path -LiteralPath $ConfigPath).Path
$config = (Get-Content -LiteralPath $resolvedConfigPath -Raw -Encoding UTF8) | ConvertFrom-Json

$gatewayHost = "127.0.0.1"
$gatewayPort = 18789
if ($null -ne $config.gateway -and $config.gateway.PSObject.Properties.Name -contains 'port' -and -not [string]::IsNullOrWhiteSpace([string]$config.gateway.port)) {
  $gatewayPort = [int]$config.gateway.port
}
if ($null -ne $config.gateway -and $config.gateway.PSObject.Properties.Name -contains 'bindHost' -and -not [string]::IsNullOrWhiteSpace([string]$config.gateway.bindHost)) {
  $gatewayHost = [string]$config.gateway.bindHost
}

$gatewayToken = $null
if ($null -ne $config.gateway -and $config.gateway.PSObject.Properties.Name -contains 'auth' -and $null -ne $config.gateway.auth -and $config.gateway.auth.PSObject.Properties.Name -contains 'token') {
  $gatewayToken = [string]$config.gateway.auth.token
}
if ([string]::IsNullOrWhiteSpace($gatewayToken)) {
  throw "OpenClaw gateway token is missing in $resolvedConfigPath"
}

$nodeCommand = Get-Command node -ErrorAction SilentlyContinue
if ($null -eq $nodeCommand -or [string]::IsNullOrWhiteSpace($nodeCommand.Source)) {
  throw "Node.js is required to reset OpenClaw sessions."
}

$nodeScript = @'
const crypto = require("crypto");
const WS = globalThis.WebSocket || require("ws");

const token = process.env.OC_RESET_GATEWAY_TOKEN;
const sessionKey = process.env.OC_RESET_SESSION_KEY;
const gatewayUrl = process.env.OC_RESET_GATEWAY_URL;

if (!token || !sessionKey || !gatewayUrl) {
  throw new Error("Missing gateway reset environment.");
}

const ws = new WS(gatewayUrl);
const pending = new Map();
let closed = false;

function request(method, params) {
  return new Promise((resolve, reject) => {
    const id = crypto.randomUUID();
    pending.set(id, { resolve, reject });
    ws.send(JSON.stringify({ type: "req", id, method, params }));
  });
}

function rejectPending(message) {
  if (closed) {
    return;
  }
  closed = true;
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
        version: "codex-reset",
        platform: process.platform,
        mode: "cli",
        instanceId: crypto.randomUUID()
      },
      role: "operator",
      scopes: ["operator.admin", "operator.approvals", "operator.pairing"],
      auth: { token },
      userAgent: "codex-reset",
      locale: "zh-CN"
    });
    const result = await request("sessions.reset", { key: sessionKey });
    process.stdout.write(JSON.stringify(result, null, 2));
    ws.close();
  } catch (error) {
    process.stderr.write((error && error.stack) ? error.stack : String(error));
    process.exitCode = 1;
    ws.close();
  }
};

ws.onmessage = (event) => {
  const raw = typeof event.data === "string" ? event.data : event.data.toString();
  const message = JSON.parse(raw);
  if (message.type !== "res") {
    return;
  }
  const slot = pending.get(message.id);
  if (!slot) {
    return;
  }
  pending.delete(message.id);
  if (message.ok) {
    slot.resolve(message.payload);
  } else {
    slot.reject(new Error((message.error && message.error.message) || "request failed"));
  }
};

ws.onerror = (error) => {
  process.stderr.write((error && error.message) ? error.message : String(error));
};

ws.onclose = () => {
  rejectPending("socket closed");
};
'@

$previousToken = $env:OC_RESET_GATEWAY_TOKEN
$previousUrl = $env:OC_RESET_GATEWAY_URL
$previousKey = $env:OC_RESET_SESSION_KEY

try {
  $env:OC_RESET_GATEWAY_TOKEN = $gatewayToken
  $env:OC_RESET_GATEWAY_URL = "ws://{0}:{1}" -f $gatewayHost, $gatewayPort
  $env:OC_RESET_SESSION_KEY = $SessionKey
  $output = $nodeScript | & $nodeCommand.Source -
  if ($LASTEXITCODE -ne 0) {
    throw "Failed to reset session key $SessionKey"
  }
  Write-Output $output
} finally {
  $env:OC_RESET_GATEWAY_TOKEN = $previousToken
  $env:OC_RESET_GATEWAY_URL = $previousUrl
  $env:OC_RESET_SESSION_KEY = $previousKey
}
