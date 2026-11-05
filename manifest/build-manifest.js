const fs = require('fs');
const path = require('path');

// Helper to parse .env file
function loadEnv(filePath) {
  if (!fs.existsSync(filePath)) return {};
  const content = fs.readFileSync(filePath, 'utf8');
  const env = {};
  for (const line of content.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const eqIdx = trimmed.indexOf('=');
    if (eqIdx !== -1) {
      const key = trimmed.slice(0, eqIdx).trim();
      const val = trimmed.slice(eqIdx + 1).trim();
      env[key] = val;
    }
  }
  return env;
}

const envPath = path.resolve(__dirname, '../server/.env');
const fileEnv = loadEnv(envPath);
const getVal = (key, fallback) => process.env[key] || fileEnv[key] || fallback;

const rawNgrokHost = getVal('NGROK_HOST', 'https://strict-huff-roast.ngrok-free.dev');
const rawNgrokDomain = getVal('NGROK_DOMAIN', rawNgrokHost.replace(/^https?:\/\//, ''));
const platformDomain = getVal('PLATFORM_DOMAIN', 'platform4x.com');
const botId = getVal('BOT_ID', getVal('MicrosoftAppId', '67127d20-25a1-4da1-9778-2d9040e6fcb5'));
const appId = getVal('TEAMS_APP_ID', 'e93f7b21-4d89-4a90-8c9f-35da124e7c19');
const devName = getVal('DEVELOPER_NAME', 'Platform4x');
const devUrl = getVal('DEVELOPER_WEBSITE', 'https://platform4x.com');

const manifest = {
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.16/MicrosoftTeams.schema.json",
  "manifestVersion": "1.16",
  "version": "1.0.0",
  "id": appId,
  "packageName": getVal('PACKAGE_NAME', 'com.platform4x.adaptivecards.v2'),
  "developer": {
    "name": devName,
    "websiteUrl": devUrl,
    "privacyUrl": `${devUrl}/privacy`,
    "termsOfUseUrl": `${devUrl}/terms`
  },
  "icons": {
    "color": "color.png",
    "outline": "outline.png"
  },
  "name": {
    "short": getVal('APP_NAME_SHORT', 'Platform4x Approvals'),
    "full": getVal('APP_NAME_FULL', 'Platform4x Adaptive Cards & Approvals Bot')
  },
  "description": {
    "short": "Interactive Adaptive Cards for Platform4x leave requests and workflows",
    "full": "Send, view and approve Platform4x leave workflows, expense claims and actionable notifications directly inside Microsoft Teams using Adaptive Cards."
  },
  "accentColor": "#4F46E5",
  "bots": [
    {
      "botId": botId,
      "scopes": ["personal", "team", "groupchat"],
      "supportsFiles": false,
      "isNotificationOnly": false
    }
  ],
  "composeExtensions": [
    {
      "botId": botId,
      "canUpdateConfiguration": false,
      "commands": [
        {
          "id": "searchWorkflows",
          "type": "query",
          "title": "Search Platform4x",
          "description": "Find and insert Platform4x workflow approval cards",
          "initialRun": true,
          "context": ["compose", "commandBox"],
          "parameters": [
            {
              "name": "query",
              "title": "Workflow Name or ID",
              "description": "Enter employee name, request ID or leave type",
              "inputType": "text"
            }
          ]
        }
      ],
      "messageHandlers": [
        {
          "type": "link",
          "value": {
            "domains": [
              platformDomain,
              `*.${platformDomain}`,
              rawNgrokDomain
            ]
          }
        }
      ]
    }
  ],
  "permissions": ["identity", "messageTeamMembers"],
  "validDomains": [
    platformDomain,
    `*.${platformDomain}`,
    "*.ngrok-free.dev",
    "*.ngrok.io",
    rawNgrokDomain
  ]
};

const outputPath = path.resolve(__dirname, 'manifest.json');
fs.writeFileSync(outputPath, JSON.stringify(manifest, null, 2), 'utf8');
console.log('Successfully generated valid manifest/manifest.json from .env variables!');
