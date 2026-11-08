// Default Seed Data
const initialRequests = [
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
    fromDate: "2026-11-15",
    toDate: "2026-11-16",
    totalDays: 2,
    reason: "Attending friend's wedding reception in Bangalore.",
    status: "Pending",
    approverComment: "",
    updatedAt: new Date().toISOString()
  }
];

// In-memory fallback
let memoryRequests = JSON.parse(JSON.stringify(initialRequests));

// Helper to get synced requests from KV or memory
async function getStoredRequests(env) {
  if (env && env.LEAVE_REQUESTS) {
    try {
      const stored = await env.LEAVE_REQUESTS.get("requests", "json");
      if (stored && Array.isArray(stored)) {
        return stored;
      }
      await env.LEAVE_REQUESTS.put("requests", JSON.stringify(initialRequests));
    } catch (e) {}
  }
  return memoryRequests;
}

// Helper to save synced requests to KV and memory
async function saveStoredRequests(env, data) {
  memoryRequests = data;
  if (env && env.LEAVE_REQUESTS) {
    try {
      await env.LEAVE_REQUESTS.put("requests", JSON.stringify(data));
    } catch (e) {}
  }
}

function buildAdaptiveCard(reqData) {
  const isApproved = reqData.status === "Approved";
  const isRejected = reqData.status === "Rejected";
  const isPending = !isApproved && !isRejected;

  const card = {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.4",
    "body": [
      {
        "type": "Container",
        "style": isApproved ? "good" : isRejected ? "attention" : "emphasis",
        "bleed": true,
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
                    "text": "Platform4x • Leave Request",
                    "size": "Small",
                    "weight": "Bolder",
                    "color": isApproved ? "Good" : isRejected ? "Attention" : "Accent"
                  },
                  {
                    "type": "TextBlock",
                    "text": reqData.employeeName,
                    "size": "Large",
                    "weight": "Bolder",
                    "spacing": "None"
                  },
                  {
                    "type": "TextBlock",
                    "text": `Request ID: ${reqData.id}`,
                    "isSubtle": true,
                    "size": "Small",
                    "spacing": "None"
                  }
                ]
              }
            ]
          }
        ]
      },
      {
        "type": "Container",
        "spacing": "Medium",
        "items": [
          {
            "type": "FactSet",
            "facts": [
              { "title": "Leave Type:", "value": reqData.leaveType },
              { "title": "Duration:", "value": `${reqData.fromDate} to ${reqData.toDate} (${reqData.totalDays} Days)` },
              { "title": "Reason:", "value": reqData.reason },
              { "title": "Status:", "value": reqData.status.toUpperCase() }
            ]
          }
        ]
      }
    ]
  };

  if (!isPending) {
    card.body.push({
      "type": "Container",
      "style": isApproved ? "good" : "attention",
      "bleed": true,
      "items": [
        {
          "type": "TextBlock",
          "text": isApproved ? "✅ DECISION RECORDED: APPROVED" : "❌ DECISION RECORDED: REJECTED",
          "weight": "Bolder",
          "color": isApproved ? "Good" : "Attention"
        },
        {
          "type": "TextBlock",
          "text": `Remarks: ${reqData.approverComment || "No remarks entered"}`,
          "isSubtle": true,
          "wrap": true
        }
      ]
    });
  } else {
    card.body.push({
      "type": "Input.Text",
      "id": "managerComment",
      "placeholder": "Add optional remarks..."
    });
    card.actions = [
      {
        "type": "Action.Execute",
        "id": "btnApprove",
        "title": "Approve ✅",
        "style": "positive",
        "verb": "APPROVE",
        "data": { "action": "APPROVE", "requestId": reqData.id }
      },
      {
        "type": "Action.Execute",
        "id": "btnReject",
        "title": "Reject ❌",
        "style": "destructive",
        "verb": "REJECT",
        "data": { "action": "REJECT", "requestId": reqData.id }
      },
      {
        "type": "Action.OpenUrl",
        "title": "Open in Platform4x ↗",
        "url": `https://platform4x.com/workflows/${reqData.id}`
      }
    ];
  }

  return card;
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const corsHeaders = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    };

    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders });
    }

    const leaveRequests = await getStoredRequests(env);

    // 1. GET /api/requests
    if (url.pathname === '/api/requests' && request.method === 'GET') {
      return new Response(JSON.stringify({ success: true, data: leaveRequests }), {
        headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      });
    }

    // 2. GET /api/card/:reqId
    if (url.pathname.startsWith('/api/card/') && request.method === 'GET') {
      const reqId = url.pathname.replace('/api/card/', '');
      const reqData = leaveRequests.find(r => r.id === reqId) || leaveRequests[0];
      return new Response(JSON.stringify({
        success: true,
        data: reqData,
        adaptiveCard: buildAdaptiveCard(reqData)
      }), {
        headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      });
    }

    // 3. POST /api/action
    if (url.pathname === '/api/action' && request.method === 'POST') {
      try {
        const body = await request.json();
        const action = body.action || "APPROVE";
        const requestId = body.requestId || "REQ-1001";
        const comment = body.managerComment || "";

        const targetReq = leaveRequests.find(r => r.id === requestId) || leaveRequests[0];
        if (action === "APPROVE") targetReq.status = "Approved";
        if (action === "REJECT") targetReq.status = "Rejected";
        targetReq.approverComment = comment;
        targetReq.updatedAt = new Date().toISOString();

        await saveStoredRequests(env, leaveRequests);

        return new Response(JSON.stringify({
          success: true,
          message: `Request ${targetReq.id} marked as ${targetReq.status}`,
          updatedRequest: targetReq,
          updatedCard: buildAdaptiveCard(targetReq)
        }), {
          headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        });
      } catch (err) {
        return new Response(JSON.stringify({ error: err.message }), {
          status: 400,
          headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        });
      }
    }

    // 4. POST /api/messages (Teams & Outlook Bot Webhook)
    if (url.pathname === '/api/messages' && request.method === 'POST') {
      try {
        const activity = await request.json();
        const activityType = activity.type;
        const activityName = activity.name;

        // Outlook / Teams Message Extension Search Query
        if (activityType === 'invoke' && activityName === 'composeExtension/query') {
          const queryText = (activity.value && activity.value.parameters && activity.value.parameters[0] && activity.value.parameters[0].value) ? activity.value.parameters[0].value.toLowerCase() : "";
          const filtered = queryText
            ? leaveRequests.filter(r => r.employeeName.toLowerCase().includes(queryText) || r.id.toLowerCase().includes(queryText) || r.leaveType.toLowerCase().includes(queryText))
            : leaveRequests;

          const attachments = (filtered.length ? filtered : leaveRequests).map(req => ({
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: buildAdaptiveCard(req),
            preview: {
              contentType: 'application/vnd.microsoft.card.thumbnail',
              content: {
                title: `${req.id}: ${req.employeeName} (${req.leaveType})`,
                subtitle: `${req.totalDays} Days • Status: ${req.status}`,
                text: req.reason,
                images: [{ url: req.employeeAvatar }]
              }
            }
          }));

          return new Response(JSON.stringify({
            composeExtension: {
              type: 'result',
              attachmentLayout: 'list',
              attachments: attachments
            }
          }), {
            headers: { ...corsHeaders, 'Content-Type': 'application/json' }
          });
        }

        // Link Unfurling Query
        if (activityType === 'invoke' && activityName === 'composeExtension/queryLink') {
          const queryUrl = (activity.value && activity.value.url) ? activity.value.url : "";
          const match = queryUrl.match(/REQ-\d+/i);
          const reqId = match ? match[0].toUpperCase() : "REQ-1001";
          const targetReq = leaveRequests.find(r => r.id === reqId) || leaveRequests[0];
          const cardJson = buildAdaptiveCard(targetReq);

          return new Response(JSON.stringify({
            composeExtension: {
              type: 'result',
              attachmentLayout: 'list',
              attachments: [
                {
                  contentType: 'application/vnd.microsoft.card.adaptive',
                  content: cardJson,
                  preview: {
                    contentType: 'application/vnd.microsoft.card.thumbnail',
                    content: {
                      title: `Platform4x: ${targetReq.employeeName}`,
                      subtitle: `${targetReq.leaveType} (${targetReq.totalDays} Days)`,
                      text: targetReq.reason,
                      images: [{ url: targetReq.employeeAvatar }]
                    }
                  }
                }
              ]
            }
          }), {
            headers: { ...corsHeaders, 'Content-Type': 'application/json' }
          });
        }

        // Universal Adaptive Card Action Invoke (Approve / Reject in Teams / Outlook)
        if (activityType === 'invoke' && (activityName === 'adaptiveCard/action')) {
          const data = (activity.value && activity.value.action && activity.value.action.data) || {};
          const action = data.action || activity.value.verb || "APPROVE";
          const requestId = data.requestId || "REQ-1001";
          const comment = (activity.value && activity.value.data && activity.value.data.managerComment) || "";

          const targetReq = leaveRequests.find(r => r.id === requestId) || leaveRequests[0];
          if (action === "APPROVE") targetReq.status = "Approved";
          if (action === "REJECT") targetReq.status = "Rejected";
          targetReq.approverComment = comment;
          targetReq.updatedAt = new Date().toISOString();

          // Persist to Cloudflare KV Database
          await saveStoredRequests(env, leaveRequests);

          return new Response(JSON.stringify({
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: buildAdaptiveCard(targetReq)
          }), {
            headers: { ...corsHeaders, 'Content-Type': 'application/json' }
          });
        }

        // Normal Chat message ("leave")
        if (activityType === 'message') {
          const targetReq = leaveRequests[0];
          const cardJson = buildAdaptiveCard(targetReq);

          const serviceUrl = activity.serviceUrl ? activity.serviceUrl.replace(/\/$/, '') : "";
          const conversationId = activity.conversation ? activity.conversation.id : "";
          const appId = env.BOT_ID || "67127d20-25a1-4da1-9778-2d9040e6fcb5";
          const secret = env.BOT_PASSWORD || "";
          const tenantId = env.APP_TENANT_ID || "917890c2-df41-411f-84c1-384442ecc59b";

          // Teams push delivery requires Bot Connector REST API call
          if (serviceUrl && conversationId) {
            try {
              const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: new URLSearchParams({
                  client_id: appId,
                  client_secret: secret,
                  grant_type: 'client_credentials',
                  scope: 'https://api.botframework.com/.default'
                })
              });
              const tokenData = await tokenRes.json();
              if (tokenData.access_token) {
                await fetch(`${serviceUrl}/v3/conversations/${conversationId}/activities`, {
                  method: 'POST',
                  headers: {
                    'Authorization': `Bearer ${tokenData.access_token}`,
                    'Content-Type': 'application/json'
                  },
                  body: JSON.stringify({
                    type: 'message',
                    attachments: [
                      {
                        contentType: 'application/vnd.microsoft.card.adaptive',
                        content: cardJson
                      }
                    ]
                  })
                });
              }
            } catch (postErr) {}
          }

          return new Response(JSON.stringify({
            type: 'message',
            attachments: [
              {
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: cardJson
              }
            ]
          }), {
            headers: { ...corsHeaders, 'Content-Type': 'application/json' }
          });
        }

        return new Response(JSON.stringify({ status: 'ok' }), {
          headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        });
      } catch (err) {
        return new Response(JSON.stringify({ error: err.message }), {
          status: 200,
          headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        });
      }
    }

    return new Response(JSON.stringify({ message: "Platform4x Adaptive Cards API is Live on Cloudflare Workers!" }), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' }
    });
  }
};
