require('dotenv').config();
const express = require('express');
const cors = require('cors');
const {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  MessageFactory,
  CardFactory,
  TeamsActivityHandler
} = require('botbuilder');

const app = express();
const PORT = 5000;

app.use(cors());
app.use(express.json());

// In-memory data store for requests
let leaveRequests = [
  {
    id: "REQ-1001",
    employeeName: "Rahul Kumar",
    employeeAvatar: "https://images.unsplash.com/photo-1535713875002-d1d0cf377fde?w=100&auto=format&fit=crop&q=80",
    leaveType: "Sick / Medical Leave",
    fromDate: "2026-11-05",
    toDate: "2026-11-07",
    totalDays: 3,
    reason: "Severe flu and fever; doctor advised 3 days complete bed rest.",
    status: "Pending",
    approverComment: "",
    updatedAt: new Date().toISOString()
  },
  {
    id: "REQ-1002",
    employeeName: "Priya Sharma",
    employeeAvatar: "https://images.unsplash.com/photo-1494790108377-be9c29b29330?w=100&auto=format&fit=crop&q=80",
    leaveType: "Casual Leave",
    fromDate: "2026-11-12",
    toDate: "2026-11-13",
    totalDays: 2,
    reason: "Attending sister's wedding ceremony in hometown.",
    status: "Pending",
    approverComment: "",
    updatedAt: new Date().toISOString()
  },
  {
    id: "REQ-1003",
    employeeName: "Rahul Sharma",
    employeeAvatar: "https://images.unsplash.com/photo-1494790108377-be9c29b29330?w=100&auto=format&fit=crop&q=80",
    leaveType: "Casual Leave",
    fromDate: "2026-11-12",
    toDate: "2026-11-13",
    totalDays: 2,
    reason: "Attending sister's wedding ceremony in hometown.",
    status: "Pending",
    approverComment: "",
    updatedAt: new Date().toISOString()
  }
];

// Build Adaptive Card JSON payload (Version 1.4 / 1.5 - Teams & Outlook compatible)
function buildAdaptiveCard(reqData, originatorId = "") {
  const isPending = reqData.status === "Pending";
  const statusColor = reqData.status === "Approved" ? "Good" : (reqData.status === "Rejected" ? "Attention" : "Warning");

  const card = {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.4",
    "body": [
      {
        "type": "Container",
        "style": "emphasis",
        "items": [
          {
            "type": "ColumnSet",
            "columns": [
              {
                "type": "Column",
                "width": "auto",
                "items": [
                  {
                    "type": "Image",
                    "url": reqData.employeeAvatar,
                    "size": "Medium",
                    "style": "Person"
                  }
                ]
              },
              {
                "type": "Column",
                "width": "stretch",
                "items": [
                  {
                    "type": "TextBlock",
                    "text": "Platform4x Leave Approval Request",
                    "weight": "Bolder",
                    "size": "Medium",
                    "wrap": true
                  },
                  {
                    "type": "TextBlock",
                    "text": `Submitted by **${reqData.employeeName}** • ID: ${reqData.id}`,
                    "isSubtle": true,
                    "spacing": "None",
                    "wrap": true
                  }
                ]
              },
              {
                "type": "Column",
                "width": "auto",
                "items": [
                  {
                    "type": "TextBlock",
                    "text": reqData.status.toUpperCase(),
                    "weight": "Bolder",
                    "color": statusColor,
                    "wrap": true
                  }
                ]
              }
            ]
          }
        ]
      },
      {
        "type": "FactSet",
        "spacing": "Medium",
        "facts": [
          { "title": "Leave Type:", "value": reqData.leaveType },
          { "title": "Duration:", "value": `${reqData.fromDate} to ${reqData.toDate} (${reqData.totalDays} Days)` },
          { "title": "Reason:", "value": reqData.reason },
          { "title": "Current Status:", "value": reqData.status }
        ]
      }
    ]
  };

  if (originatorId) {
    card.originator = originatorId;
  }

  if (!isPending) {
    card.body.push({
      "type": "Container",
      "style": reqData.status === "Approved" ? "good" : "attention",
      "spacing": "Medium",
      "items": [
        {
          "type": "TextBlock",
          "text": `**Decision Recorded:** Request was **${reqData.status}**.`,
          "wrap": true
        },
        {
          "type": "TextBlock",
          "text": reqData.approverComment ? `Remarks: "${reqData.approverComment}"` : "No remarks provided.",
          "isSubtle": true,
          "wrap": true
        }
      ]
    });
  } else {
    card.body.push({
      "type": "Input.Text",
      "id": "managerComment",
      "placeholder": "Add optional notes or approval reason...",
      "isMultiline": true,
      "spacing": "Medium"
    });

    card.actions = [
      {
        "type": "Action.Execute",
        "title": "Approve ✅",
        "verb": "approve",
        "style": "positive",
        "data": {
          "action": "APPROVE",
          "requestId": reqData.id
        }
      },
      {
        "type": "Action.Execute",
        "title": "Reject ❌",
        "verb": "reject",
        "style": "destructive",
        "data": {
          "action": "REJECT",
          "requestId": reqData.id
        }
      },
      {
        "type": "Action.OpenUrl",
        "title": "Open in Platform4x ↗",
        "url": `http://localhost:3000/?reqId=${reqData.id}`
      }
    ];
  }

  return card;
}

// -------------------------------------------------------------
// TEAMS BOT INTEGRATION (OFFICIAL SDK HANDLER)
// -------------------------------------------------------------
const botFrameworkAuth = new ConfigurationBotFrameworkAuthentication({
  MicrosoftAppId: process.env.MicrosoftAppId,
  MicrosoftAppPassword: process.env.MicrosoftAppPassword,
  MicrosoftAppType: process.env.MicrosoftAppType || "MultiTenant",
  MicrosoftAppTenantId: process.env.MicrosoftAppTenantId
});

const adapter = new CloudAdapter(botFrameworkAuth);

// Helper to fetch direct bot connector token for single-tenant Teams delivery
async function getBotConnectorToken() {
  const appId = process.env.MicrosoftAppId;
  const secret = process.env.MicrosoftAppPassword;
  const tenantId = process.env.MicrosoftAppTenantId;

  const res = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    body: new URLSearchParams({
      client_id: appId,
      client_secret: secret,
      grant_type: 'client_credentials',
      scope: 'https://api.botframework.com/.default'
    })
  });
  const data = await res.json();
  return data.access_token;
}

adapter.onTurnError = async (context, error) => {
  // If Teams SDK sends response with tenant mismatch, deliver directly via Bot Connector REST API
  if (error.message && error.message.includes('Authorization has been denied')) {
    try {
      const token = await getBotConnectorToken();
      const serviceUrl = context.activity.serviceUrl;
      const conversationId = context.activity.conversation.id;
      const targetReq = leaveRequests[0];
      const cardJson = buildAdaptiveCard(targetReq);

      await fetch(`${serviceUrl}v3/conversations/${conversationId}/activities`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          type: 'message',
          attachments: [CardFactory.adaptiveCard(cardJson)]
        })
      });
      console.log(`[Auto-Delivered] Direct connector message delivered to Teams successfully!`);
      return;
    } catch (e) {
      console.error(`Fallback delivery error:`, e.message);
    }
  }
  console.error(`[onTurnError] Bot error: ${error.message || error}`);
};

class TeamsApprovalBot extends TeamsActivityHandler {
  constructor() {
    super();

    // 1. Text Message Handler ("leave", "hello")
    this.onMessage(async (context, next) => {
      const targetReq = leaveRequests[0];
      const cardJson = buildAdaptiveCard(targetReq);
      const cardActivity = MessageFactory.attachment(CardFactory.adaptiveCard(cardJson));
      await context.sendActivity(cardActivity);
      await next();
    });
  }

  // 2. Official Universal Action / In-Place Adaptive Card Update Handler
  async onAdaptiveCardInvoke(context, invokeValue) {
    console.log(`>>> [AdaptiveCard Invoke Handler] Action verb: ${invokeValue.action?.verb || 'action'}`);
    const data = invokeValue.action?.data || {};
    const { action, requestId } = data;
    const comment = invokeValue.action?.data?.managerComment || "";

    const targetReq = leaveRequests.find(r => r.id === requestId) || leaveRequests[0];
    if (action === "APPROVE") targetReq.status = "Approved";
    if (action === "REJECT") targetReq.status = "Rejected";
    targetReq.approverComment = comment;

    const updatedCard = buildAdaptiveCard(targetReq);

    // Official Microsoft Invoke Response
    return {
      statusCode: 200,
      type: 'application/vnd.microsoft.card.adaptive',
      value: updatedCard
    };
  }

  // 3. Official Message Extension Query (+ button / search in Teams & Outlook)
  async handleTeamsMessagingExtensionQuery(context, query) {
    const searchQuery = (query.parameters && query.parameters[0] && query.parameters[0].value) || '';
    console.log(`>>> [Message Extension Query] Searching for: "${searchQuery}"`);

    const filtered = leaveRequests.filter(r =>
      r.employeeName.toLowerCase().includes(searchQuery.toLowerCase()) ||
      r.id.toLowerCase().includes(searchQuery.toLowerCase()) ||
      r.leaveType.toLowerCase().includes(searchQuery.toLowerCase())
    );

    const attachments = (filtered.length > 0 ? filtered : leaveRequests).map(req => {
      const cardJson = buildAdaptiveCard(req);
      const previewCard = CardFactory.thumbnailCard(
        `${req.employeeName} (${req.id})`,
        `${req.leaveType} - ${req.status}`,
        [{ url: req.employeeAvatar }]
      );
      return {
        ...CardFactory.adaptiveCard(cardJson),
        preview: previewCard
      };
    });

    return {
      composeExtension: {
        type: 'result',
        attachmentLayout: 'list',
        attachments: attachments
      }
    };
  }

  // 4. Official Link Unfurling Handler (When link is pasted in Teams chat)
  async handleTeamsAppBasedLinkQuery(context, query) {
    console.log(`>>> [Link Unfurling Triggered] URL: ${query.url}`);

    // Extract Request ID from URL if present (e.g., REQ-1001 or REQ-1002)
    const match = (query.url || '').match(/REQ-\d+/i);
    const reqId = match ? match[0].toUpperCase() : 'REQ-1001';
    const targetReq = leaveRequests.find(r => r.id === reqId) || leaveRequests[0];

    const cardJson = buildAdaptiveCard(targetReq);
    const previewCard = CardFactory.thumbnailCard(
      `Platform4x: ${targetReq.employeeName} (${targetReq.id})`,
      `${targetReq.leaveType} (${targetReq.totalDays} Days) - ${targetReq.status}`,
      [{ url: targetReq.employeeAvatar }]
    );

    return {
      composeExtension: {
        type: 'result',
        attachmentLayout: 'list',
        attachments: [
          {
            ...CardFactory.adaptiveCard(cardJson),
            preview: previewCard
          }
        ]
      }
    };
  }
}

const bot = new TeamsApprovalBot();

// Unified endpoint for Teams & Azure Bot WebChat
app.post('/api/messages', async (req, res) => {
  // If request comes from Outlook Message Extension (channelId: m365extensions)
  // handle it directly to bypass BotFramework SDK's internal signing key caching latency
  if (req.body && (req.body.channelId === 'm365extensions' || req.headers['channelid'] === 'm365extensions')) {
    const activity = req.body;
    console.log(`>>> [Outlook M365 Channel Detected] Activity: ${activity.name || activity.type}`);

    if (activity.name === 'composeExtension/query') {
      const query = activity.value || {};
      const searchQuery = (query.parameters && query.parameters[0] && query.parameters[0].value) || '';
      console.log(`>>> [Outlook Query] Search term: "${searchQuery}"`);

      let results = leaveRequests;
      if (searchQuery && searchQuery.trim() !== '' && searchQuery !== 'true') {
        const q = searchQuery.toLowerCase();
        results = leaveRequests.filter(r =>
          r.employeeName.toLowerCase().includes(q) ||
          r.id.toLowerCase().includes(q) ||
          r.leaveType.toLowerCase().includes(q)
        );
      }

      const attachments = results.map(r => {
        const cardJson = buildAdaptiveCard(r);
        return {
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: cardJson,
          preview: {
            contentType: 'application/vnd.microsoft.card.hero',
            content: {
              title: `${r.employeeName} (${r.id})`,
              subtitle: `${r.leaveType} • ${r.totalDays} Days • Status: ${r.status}`,
              text: r.reason
            }
          }
        };
      });

      return res.status(200).json({
        composeExtension: {
          type: 'result',
          attachmentLayout: 'list',
          attachments: attachments
        }
      });
    }

    if (activity.name === 'composeExtension/queryLink') {
      const url = activity.value?.url || '';
      console.log(`>>> [Outlook Link Unfurling] URL: ${url}`);
      const targetReq = url.includes('REQ-1002') ? leaveRequests[1] : leaveRequests[0];
      const cardJson = buildAdaptiveCard(targetReq);

      return res.status(200).json({
        composeExtension: {
          type: 'result',
          attachmentLayout: 'list',
          attachments: [
            {
              contentType: 'application/vnd.microsoft.card.adaptive',
              content: cardJson,
              preview: {
                contentType: 'application/vnd.microsoft.card.hero',
                content: {
                  title: `${targetReq.employeeName} (${targetReq.id})`,
                  subtitle: `${targetReq.leaveType} • Status: ${targetReq.status}`,
                  text: targetReq.reason
                }
              }
            }
          ]
        }
      });
    }

    // Handle in-place Adaptive Card Action from Outlook email!
    if (activity.name === 'adaptiveCard/action') {
      const data = activity.value?.action?.data || {};
      const { action, requestId } = data;
      const comment = data.managerComment || "";
      console.log(`>>> [Outlook In-Place Action] Action: ${action}, RequestId: ${requestId}`);

      const targetReq = leaveRequests.find(r => r.id === requestId) || leaveRequests[0];
      if (action === "APPROVE") targetReq.status = "Approved";
      if (action === "REJECT") targetReq.status = "Rejected";
      targetReq.approverComment = comment;
      targetReq.updatedAt = new Date().toISOString();

      const updatedCard = buildAdaptiveCard(targetReq);

      return res.status(200).json({
        statusCode: 200,
        type: 'application/vnd.microsoft.card.adaptive',
        value: updatedCard
      });
    }
  }

  try {
    await adapter.process(req, res, async (context) => {
      await bot.run(context);
    });
  } catch (err) {
    console.error('Error processing bot activity via CloudAdapter:', err.message);
    res.status(200).send();
  }
});

// -------------------------------------------------------------
// OUTLOOK ACTIONABLE EMAIL SENDER API
// -------------------------------------------------------------
app.post('/api/send-actionable-email', async (req, res) => {
  const { toEmail, requestId = "REQ-1001" } = req.body;
  if (!toEmail) {
    return res.status(400).json({ error: "toEmail is required" });
  }

  const targetReq = leaveRequests.find(r => r.id === requestId) || leaveRequests[0];
  const ngrokHost = process.env.NGROK_HOST || "https://strict-huff-roast.ngrok-free.dev";

  // Build the Outlook Actionable Adaptive Card (Version 1.4 strictly supported by Outlook)
  const outlookCard = {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.4",
    "body": [
      {
        "type": "Container",
        "style": "emphasis",
        "items": [
          {
            "type": "ColumnSet",
            "columns": [
              {
                "type": "Column",
                "width": "auto",
                "items": [
                  {
                    "type": "Image",
                    "url": targetReq.employeeAvatar,
                    "size": "Medium",
                    "style": "Person"
                  }
                ]
              },
              {
                "type": "Column",
                "width": "stretch",
                "items": [
                  {
                    "type": "TextBlock",
                    "text": "Platform4x Leave Approval Request",
                    "weight": "Bolder",
                    "size": "Medium"
                  },
                  {
                    "type": "TextBlock",
                    "text": `Submitted by **${targetReq.employeeName}** • ID: ${targetReq.id}`,
                    "isSubtle": true,
                    "spacing": "None"
                  }
                ]
              }
            ]
          }
        ]
      },
      {
        "type": "FactSet",
        "facts": [
          { "title": "Leave Type:", "value": targetReq.leaveType },
          { "title": "Duration:", "value": `${targetReq.fromDate} to ${targetReq.toDate} (${targetReq.totalDays} Days)` },
          { "title": "Reason:", "value": targetReq.reason },
          { "title": "Current Status:", "value": targetReq.status }
        ]
      },
      {
        "type": "Input.Text",
        "id": "managerComment",
        "placeholder": "Add optional remarks..."
      }
    ],
    "actions": [
      {
        "type": "Action.Http",
        "title": "Approve ✅",
        "method": "POST",
        "url": `${ngrokHost}/api/action`,
        "body": JSON.stringify({
          action: "APPROVE",
          requestId: targetReq.id,
          managerComment: "{{managerComment.value}}"
        }),
        "headers": [
          { "name": "Content-Type", "value": "application/json" }
        ]
      },
      {
        "type": "Action.Http",
        "title": "Reject ❌",
        "method": "POST",
        "url": `${ngrokHost}/api/action`,
        "body": JSON.stringify({
          action: "REJECT",
          requestId: targetReq.id,
          managerComment: "{{managerComment.value}}"
        }),
        "headers": [
          { "name": "Content-Type", "value": "application/json" }
        ]
      },
      {
        "type": "Action.OpenUrl",
        "title": "Open in Platform4x ↗",
        "url": `http://localhost:3000/?reqId=${targetReq.id}`
      }
    ]
  };

  // Full Actionable Email HTML wrapper with embedded JSON
  const actionableEmailHtml = `
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <script type="application/adaptivecard+json">
  ${JSON.stringify(outlookCard, null, 2)}
  <\/script>
</head>
<body>
  <div style="font-family: Segoe UI, sans-serif; padding: 15px; color: #333;">
    <h3>Platform4x Leave Approval Request</h3>
    <p>Please view this email in Outlook to approve or reject directly from your inbox.</p>
    <p><strong>Employee:</strong> ${targetReq.employeeName}<br/>
       <strong>Leave:</strong> ${targetReq.leaveType} (${targetReq.totalDays} Days)<br/>
       <strong>Reason:</strong> ${targetReq.reason}</p>
    <a href="http://localhost:3000/?reqId=${targetReq.id}" style="color: #4f46e5; font-weight: bold;">Review in Platform4x</a>
  </div>
</body>
</html>
  `.trim();

  res.json({
    success: true,
    message: "Actionable Message HTML payload generated successfully",
    targetEmail: toEmail,
    adaptiveCard: outlookCard,
    embeddableHtml: actionableEmailHtml
  });
});
app.get('/api/requests', (req, res) => {
  res.json({ success: true, data: leaveRequests });
});

app.get('/api/card/:id', (req, res) => {
  const reqData = leaveRequests.find(r => r.id === req.params.id);
  if (!reqData) return res.status(404).json({ error: "Request not found" });
  res.json({ success: true, data: reqData, adaptiveCard: buildAdaptiveCard(reqData) });
});

// Unified Action endpoint (Used by Outlook Actionable Messages & Web)
app.post('/api/action', (req, res) => {
  const { action, requestId, managerComment } = req.body;
  const targetReq = leaveRequests.find(r => r.id === requestId);
  if (!targetReq) return res.status(404).json({ error: "Request not found" });

  if (action === "APPROVE") targetReq.status = "Approved";
  if (action === "REJECT") targetReq.status = "Rejected";
  targetReq.approverComment = managerComment || "";
  targetReq.updatedAt = new Date().toISOString();

  const updatedCard = buildAdaptiveCard(targetReq);

  // Set CARD-ACTION-STATUS header for Outlook Actionable Messages instant feedback
  res.setHeader('CARD-ACTION-STATUS', `Leave request ${targetReq.id} marked as ${targetReq.status}!`);
  res.json({
    success: true,
    message: `Request ${requestId} successfully updated to ${targetReq.status}`,
    data: targetReq,
    adaptiveCard: updatedCard
  });
});

app.listen(PORT, () => {
  console.log(`Backend Server running with Teams Bot & Outlook Action Handlers on http://localhost:${PORT}`);
});
