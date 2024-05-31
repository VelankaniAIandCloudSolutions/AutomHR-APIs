const mysql = require("mysql2/promise");
const nodemailer = require("nodemailer");
const moment = require("moment");

// Database connection configuration
const dbConfig = {
  host: "localhost",
  user: "root",
  password: "Velankanidb@2123",
  database: "marx_db",
};

// Email configuration
const transporter = nodemailer.createTransport({
  service: "gmail", // or your email provider
  auth: {
    user: "your-email@example.com",
    pass: "your-email-password",
  },
});

// Function to send an email
async function sendEmail(to, subject, text) {
  let mailOptions = {
    from: "your-email@example.com",
    to,
    subject,
    text,
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log(`Email sent to ${to}`);
  } catch (error) {
    console.error(`Error sending email to ${to}:`, error);
  }
}

// Main function to check timesheets
async function checkTimesheets() {
  const connection = await mysql.createConnection(dbConfig);

  try {
    const [users] = await connection.execute("SELECT * FROM tblstaff");
    const currentMonth = moment().format("YYYY-MM");
    console.log(users[0])
    for (let user of users) {
      const [projects] = await connection.execute(
        "SELECT * FROM tblproject_members WHERE staff_id = ?",
        [user.staffid]
      );

      for (let project of projects) {
        const [tasks] = await connection.execute(
          "SELECT t.* FROM tbltasks t " +
            "INNER JOIN tbltask_assigned ta ON t.id = ta.task_id " +
            "WHERE t.rel_id = ? AND t.rel_type = ? AND ta.staffid = ?",
          [project.id, "project", user.staffid]
        );

        let hasActiveTimer = false;

        for (let task of tasks) {
          const [taskTimers] = await connection.execute(
            `SELECT * FROM tbltaskstimers 
                WHERE task_id = ? 
                AND DATE_FORMAT(FROM_UNIXTIME(start_time), '%Y-%m') = ?
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
            `SELECT * FROM tbltime_sheet_approval WHERE staff_id = ? AND project_ids = ? AND DATE_FORMAT(created_at, '%Y-%m') = ? AND status != 2`,
            [user.id, project.id, currentMonth]
          );

          if (timesheet.length === 0) {
            console.log(
              `Timesheet missing for user ${user.firstname} ${user.lastname} on project ${project.name} for the month of ${currentMonth}`
            );
            // await sendEmail(
            //   user.email,
            //   "Timesheet Missing",
            //   `Please submit your timesheet for the project ${project.name} for the month of ${currentMonth}.`
            // );
          }
        }
      }
    }
  } catch (error) {
    console.error("Error checking timesheets:", error);
  } finally {
    await connection.end();
  }
}

// Run the script
checkTimesheets();
