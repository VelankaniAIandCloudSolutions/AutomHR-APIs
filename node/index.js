const mysql = require("mysql2/promise");
const nodemailer = require("nodemailer");
const moment = require("moment");
const cron = require("node-cron");

const dbConfig = {
  host: "192.168.10.200",
  user: "ubuntu",
  password: "Velankanidb@2123",
  database: "marx_db",
};
const connection = mysql.createConnection(dbConfig);

const transporter = nodemailer.createTransport({
  host: "smtpout.secureserver.net",
  port: 465,
  secure: true,
  auth: {
    user: "info@automhr.com",
    pass: "Hotel@123!",
  },
});

async function sendEmail(to, subject, html) {
  let mailOptions = {
    from: "Marx HR <info@automhr.com>",
    to,
    subject,
    html,
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log(`Email sent to ${to}`);
  } catch (error) {
    console.error(`Error sending email to ${to}:`, error);
  }
}

async function sendLogInReminder() {
  const connection = await mysql.createConnection(dbConfig);
  try {
    const [users] = await connection.execute(
      "SELECT * FROM tblstaff WHERE status_work = ? AND staff_identifi NOT IN (?)",
      [
        "working",
        [
          "ME092",
          "ME048",
          "ME054",
          "ME055",
          "ME070",
          "ME071",
          "ME072",
          "ME079",
          "ME087",
          "ME089",
        ],
      ]
    );
    // const users = [{ email: "sharmaps112000@gmail.com", firstname: "Pranav" }];
    const subject = "Reminder: Log in to AutomHR";

    for (const user of users) {
      const html = `
        <html>
        <head>
          <style>
            .email-container {
              font-family: Arial, sans-serif;
              line-height: 1.6;
              color: #333;
            }
            .email-header {
              background-color: #f7f7f7;
              padding: 10px;
              border-bottom: 1px solid #ddd;
            }
            .email-body {
              padding: 20px;
            }
            .email-footer {
              background-color: #f7f7f7;
              padding: 10px;
              border-top: 1px solid #ddd;
              text-align: center;
              font-size: 0.9em;
            }
            .email-button {
              display: inline-block;
              padding: 10px 15px;
              margin-top: 10px;
              color: #fff;
              background-color: #007bff;
              text-decoration: none;
              border-radius: 5px;
            }
          </style>
        </head>
        <body>
          <div class="email-container">
            <div class="email-header">
              <h2>Reminder: Log in to AutomHR</h2>
            </div>
            <div class="email-body">
              <p>Dear ${user?.firstname},</p>
              <p>This is a friendly reminder to log in to <strong>AutomHR</strong> to complete your timesheet and update any pending tasks.</p>
              <p>Please make sure to log in at your earliest convenience.</p>
              <a href="https://marx.automhr.com/admin/" class="email-button">Log in to AutomHR</a>
            </div>
            <div class="email-footer">
              <p>Best regards,<br>AutomHR</p>
            </div>
          </div>
        </body>
        </html>
      `;
      // await sendEmail(user.email, subject, html);
    }
  } catch (error) {
    console.error("Error sending login reminders:", error);
  } finally {
    await connection.end();
  }
}

async function checkTimesheets() {
  const currentDate = moment();
  let currentMonth;

  if (currentDate.date() > currentDate.daysInMonth() - 2) {
    currentMonth = currentDate.format("YYYY-MM");
  } else if (currentDate.date() <= 2) {
    currentMonth = currentDate.subtract(1, "month").format("YYYY-MM");
  } else {
    console.log("Not within the specified date range for timesheet checking.");
    return;
  }
  console.log("Checking timesheets for:", currentMonth);

  const connection = await mysql.createConnection(dbConfig);

  try {
    const [users] = await connection.execute(
      "SELECT * FROM tblstaff WHERE status_work = ?",
      ["working"]
    );
    // const [users] = await connection.execute("SELECT * FROM tblstaff WHERE status_work = ? AND staffid = ?",["working",67]);

    await Promise.all(
      users.map(async (user) => {
        const [projects] = await connection.execute(
          `
        SELECT pm.*, p.name AS project_name
        FROM tblproject_members pm
        INNER JOIN tblprojects p ON pm.project_id = p.id
        WHERE pm.staff_id = ?
      `,
          [user.staffid]
        );

        await Promise.all(
          projects.map(async (project) => {
            const [tasks] = await connection.execute(
              `
          SELECT t.*, ta.staffid AS assigned_staff_id
          FROM tbltasks t
          INNER JOIN tbltask_assigned ta ON t.id = ta.taskid
          WHERE t.rel_id = ? AND t.rel_type = ? AND ta.staffid = ?
        `,
              [project.project_id, "project", user.staffid]
            );

            let hasActiveTimer = false;
            for (let task of tasks) {
              const [taskTimers] = await connection.execute(
                `
            SELECT 1
            FROM tbltaskstimers
            WHERE task_id = ? AND DATE_FORMAT(FROM_UNIXTIME(start_time), '%Y-%m') = ?
            LIMIT 1
          `,
                [task.id, currentMonth]
              );

              if (taskTimers.length > 0) {
                hasActiveTimer = true;
                break;
              }
            }

            if (tasks.length > 0 && hasActiveTimer) {
              const [timesheet] = await connection.execute(
                `
            SELECT 1
            FROM tbltime_sheet_approval
            WHERE staff_id = ? AND project_ids = ? AND DATE_FORMAT(created_at, '%Y-%m') = ? AND status != "2" AND teamlead_status != "2"
            LIMIT 1
          `,
                [user.staffid, project.project_id, currentMonth]
              );

              if (timesheet.length === 0) {
                const subject = `Missing Timesheet for the project ${
                  project.project_name
                } for the month of ${currentDate.format("MMMM, YYYY")}`;
                const html = `
              <html>
              <head>
                <style>
                  .email-container {
                    font-family: Arial, sans-serif;
                    line-height: 1.6;
                    color: #333;
                  }
                  .email-header {
                    background-color: #f7f7f7;
                    padding: 10px;
                    border-bottom: 1px solid #ddd;
                  }
                  .email-body {
                    padding: 20px;
                  }
                  .email-footer {
                    background-color: #f7f7f7;
                    padding: 10px;
                    border-top: 1px solid #ddd;
                    text-align: center;
                    font-size: 0.9em;
                  }
                  .email-button {
                    display: inline-block;
                    padding: 10px 15px;
                    margin-top: 10px;
                    color: #fff;
                    background-color: #007bff;
                    text-decoration: none;
                    border-radius: 5px;
                  }
                </style>
              </head>

              <body>
                <div class="email-container">
                  <div class="email-header">
                    <h2>Timesheet Submission Reminder</h2>
                  </div>
                  <div class="email-body">
                    <p>Dear ${user?.firstname},</p>
                    <p>This is a friendly reminder to submit your Timesheet for the project <strong> ${
                      project.project_name
                    } </strong> for the month of <strong>${currentDate.format(
                  "MMMM, YYYY"
                )}</strong> on AutomHR.</p>
                    <p>Kindly follow the link below to submit the timesheet.</p>
                    <p>Please ignore if all timesheets are submitted and approved.</p>
                    <a href="https://marx.automhr.com/admin/staff/timesheets?view=all class="email-button">Submit Timesheet</a>
                  </div>
                  <div class="email-footer">
                    <p>Best regards,<br>Marx HR</p>
                  </div>
                </div>
              </body>
              </html>
            `;
                await sendEmail(user.email, subject, html);
              }
            }
          })
        );
      })
    );
  } catch (error) {
    console.error("Error checking timesheets:", error);
  } finally {
    await connection.end();
  }
}

// cron.schedule("00 12 * * *", async () => {
//   const today = moment().format("YYYY-MM-DD");
//   const start = moment("2024-05-03");
//   const end = moment("2024-06-07");

//   if (moment(today).isBetween(start, end, "days", "[]")) {
//     await sendLogInReminder();
//   }
// });

// cron.schedule('0 12 * * *', checkTimesheets);
console.log("Scheduler is running...");
