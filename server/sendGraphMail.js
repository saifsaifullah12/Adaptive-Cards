require('dotenv').config();

async function startDeviceLogin() {
  const appId = process.env.MicrosoftAppId;
  const tenantId = process.env.MicrosoftAppTenantId;
  const targetEmail = process.env.TARGET_EMAIL || "S.Saifullah@faaztechsolutions.com";
  const ngrokHost = process.env.NGROK_HOST || "https://strict-huff-roast.ngrok-free.dev";

  console.log("------------------------------------------------------------");
  console.log("   STEP 1: Requesting 1-Click Outlook Authentication Code   ");
  console.log("------------------------------------------------------------");

  const res = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/devicecode`, {
    method: 'POST',
    body: new URLSearchParams({
      client_id: appId,
      scope: 'https://graph.microsoft.com/Mail.Send openid profile email'
    })
  });

  const data = await res.json();
  if (!res.ok) {
    console.error("Device code error:", data);
    return;
  }

  console.log("\n🔑 CODE:", data.user_code);
  console.log("👉 URL: https://microsoft.com/devicelogin");
  console.log("\nBrowser-la https://microsoft.com/devicelogin open panni mela irukira CODE-ai enter pannunga.");
  console.log("Waiting for your approval in browser...\n");

  // Poll until user signs in
  const interval = (data.interval || 5) * 1000;
  const expiresAt = Date.now() + (data.expires_in * 1000);

  while (Date.now() < expiresAt) {
    await new Promise(r => setTimeout(r, interval));

    const pollRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: 'POST',
      body: new URLSearchParams({
        grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
        client_id: appId,
        device_code: data.device_code
      })
    });

    const pollData = await pollRes.json();

    if (pollData.access_token) {
      console.log(" Authentication Successful! Access Token acquired.");
      await sendActionableCardEmail(pollData.access_token, targetEmail, ngrokHost);
      break;
    }

    if (pollData.error && pollData.error !== 'authorization_pending') {
      console.error("Authentication Error:", pollData.error_description || pollData.error);
      break;
    }
  }
}

async function sendActionableCardEmail(accessToken, targetEmail, ngrokHost) {
  console.log("\n------------------------------------------------------------");
  console.log("   STEP 2: Delivering Actionable Adaptive Card to Outlook   ");
  console.log("------------------------------------------------------------");

  const cardJson = {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.0",
    "hideOriginalBody": true,
    "body": [
      {
        "type": "Container",
        "style": "emphasis",
        "items": [
          {
            "type": "TextBlock",
            "text": "Platform4x Leave Approval Request",
            "weight": "Bolder",
            "size": "Medium",
            "color": "Dark"
          },
          {
            "type": "TextBlock",
            "text": "Submitted by Rahul Kumar • ID: REQ-1001",
            "isSubtle": true,
            "spacing": "None"
          }
        ]
      },
      {
        "type": "FactSet",
        "facts": [
          { "title": "Employee:", "value": "Rahul Kumar (Emp ID: P4X-8821)" },
          { "title": "Leave Type:", "value": "Sick / Medical Leave" },
          { "title": "Duration:", "value": "2026-11-05 to 2026-11-07 (3 Days)" },
          { "title": "Reason:", "value": "Severe flu and fever; doctor advised 3 days bed rest." },
          { "title": "Current Status:", "value": "Pending Manager Review" }
        ]
      },
      {
        "type": "Input.Text",
        "id": "managerComment",
        "placeholder": "Add optional notes or approval reason...",
        "isMultiline": true
      }
    ],
    "actions": [
      {
        "type": "Action.Http",
        "title": "Approve ✅",
        "method": "POST",
        "url": `${ngrokHost}/api/action`,
        "body": "{\"action\":\"APPROVE\",\"requestId\":\"REQ-1001\",\"managerComment\":\"{{managerComment.value}}\"}",
        "headers": [
          { "name": "Content-Type", "value": "application/json" }
        ]
      },
      {
        "type": "Action.Http",
        "title": "Reject ❌",
        "method": "POST",
        "url": `${ngrokHost}/api/action`,
        "body": "{\"action\":\"REJECT\",\"requestId\":\"REQ-1001\",\"managerComment\":\"{{managerComment.value}}\"}",
        "headers": [
          { "name": "Content-Type", "value": "application/json" }
        ]
      }
    ]
  };

  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <script type="application/adaptivecard+json">
${JSON.stringify(cardJson, null, 2)}
  </script>
</head>
<body>
  <div style="font-family: Segoe UI, sans-serif; padding: 20px;">
    <h2>Platform4x Leave Approval - REQ-1001</h2>
    <p>Please review Rahul Kumar's leave request directly in Outlook.</p>
  </div>
</body>
</html>
  `.trim();

  const sendMailRes = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      message: {
        subject: "Platform4x Actionable Leave Approval - REQ-1001",
        body: {
          contentType: "HTML",
          content: htmlBody
        },
        toRecipients: [
          {
            emailAddress: {
              address: targetEmail
            }
          }
        ]
      },
      saveToSentItems: "true"
    })
  });

  if (sendMailRes.status === 202 || sendMailRes.ok) {
    console.log("\n🎉 SUCCESS! Actionable Adaptive Card has been sent to your Outlook Inbox!");
    console.log("👉 Ippa ungaloda Outlook Inbox open panni email-ai check pannunga!");
  } else {
    console.error("❌ Send Mail Error:", sendMailRes.status, await sendMailRes.text());
  }
}

startDeviceLogin();
