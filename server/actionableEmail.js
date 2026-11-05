require('dotenv').config();
const nodemailer = require('nodemailer');

// Set your Office 365 or Outlook credentials, or test with SMTP
async function sendActionableEmail(targetEmail) {
  const ngrokHost = "https://strict-huff-roast.ngrok-free.dev";

  const cardJson = {
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
                    "text": "Submitted by **Rahul Kumar** • ID: REQ-1001",
                    "isSubtle": true
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
          { "title": "Leave Type:", "value": "Sick / Medical Leave" },
          { "title": "Duration:", "value": "2026-11-05 to 2026-11-07 (3 Days)" },
          { "title": "Reason:", "value": "Severe flu and fever; doctor advised 3 days bed rest." },
          { "title": "Current Status:", "value": "Pending" }
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

  const htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <script type="application/adaptivecard+json">
  ${JSON.stringify(cardJson, null, 2)}
  <\/script>
</head>
<body>
  <div style="font-family: Arial, sans-serif; padding: 15px;">
    <h2>Platform4x Leave Approval Request</h2>
    <p>Please take action on this leave request directly in Outlook.</p>
  </div>
</body>
</html>
  `.trim();

  console.log("--------------------------------------------------");
  console.log("Actionable Message Email Body Ready for:", targetEmail);
  console.log("--------------------------------------------------");
  return htmlContent;
}

module.exports = { sendActionableEmail };

