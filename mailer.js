const Mailer = (subject) => {
  const { clear } = require("console");
  const nodemailer = require("nodemailer");
  var emailAddress = require("./mails.json");
  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
      user: "ieeesbjec@gmail.com",
      pass: "dxvsmcpkbcqmcqvb",
    },
  });
  //Email addressess with name in dictionary structure
  // let emailAddress = {
  //   sk: "sanjaikumaran0311@gmail.com",
  // };
  let names = Object.keys(emailAddress);

  let a = 151,
    b = 1,
    flag = true;
  const sendMail = () => {
    let interval = setInterval(() => {
      if (a < names.length) {
        if (flag) {
          flag = false;
          //Body of the mail
          mailOptions[
            "html"
          ] = `<div>  <h2>Dear ${names[a]},</h2>  <h3>Greetings from IEEE SB JEC!</h3>  <p>    We are excited to present    <strong> "Codemania: The Ultimate Code Battle"</strong> an online coding    competition organized by IEEE SB JEC. Join us on    <strong>22nd July 2023, from 7:00 pm to 8:00 pm IST on HackerRank.</strong>  </p>  <img    src="https://digitalfilmartgallery.com/IEEE%20posters/sps.jpg"    alt="Codemania: The Ultimate Code Battle" />  <p>    Challenge yourself, showcase your skills, and win exciting prizes.    E-certificates will be provided to the winners.  </p>  <p>    Register here:    <a href="http://bit.ly/CodemaniaIEEE">http://bit.ly/CodemaniaIEEE</a><br />    Perks : E- certificate provided for the participants  </p>  <p>    For more information, Join the WhatsApp group :    <a href="https://chat.whatsapp.com/CRwfEOphWa81FHPBlEEXPN"></a>    https://chat.whatsapp.com/CRwfEOphWa81FHPBlEEXPN  </p>  <p>    Don't miss this opportunity to be a part of    <strong>"Codemania: The Ultimate Code Battle."</strong>  </p>  <h4>Regards,</h4>  <h4>    <strong>Renuga S</strong> <br />    Chairperson - SPS,<br />    IEEE SB JEC<br />  </h4>  <p>    For any queries contact: <br /><a href="https://wa.me/+919360480003"      >Renuga Suresh</a    ><br />    <a href="tel:9360480003">9360480003</a>  </p>  <p>    ==================================================<br />Thought of the    day<br /><br /><strong      >"Code is the poetry of a better world." - Shawn Lukas    </strong>  </p></div>`;
          mailOptions["to"] = emailAddress[names[a]];
          //To send any certificate to participants

          // mailOptions["attachments"][0]["filename"] = names[a] + ".png";
          // mailOptions["attachments"][0][
          //   "path"
          // ] = `C:/Users/DELL/Documents/Adobe/Data Set ${b}-01.png`;

          // if (job == "certificate") {
          //   attachment();
          // } else {
          //   // mailOptions["attachments"][0]["filename"] = poster + ".png";
          //   // mailOptions["attachments"][0]["path"] = poster + ".png";
          // }

          transporter.sendMail(mailOptions, function (error, info) {
            if (error) {
              //If any error occur this will execute to show the error
              mailOptions["to"] = "sanjaikumaran0311@gmail.com";
              subject = "Error occured!!";
              mailOptions[
                "html"
              ] = `<div>  <h2>Dear Sk,</h2><p>Something error occured in mail automation</p><p>${error} at ${
                a + 1
              }, ${
                names[a]
              }.<div>Time:${new Date().toLocaleString()}</div></div>`;
              transporter.sendMail(mailOptions, function (error, info) {
                if (error) {
                  console.log(error);
                }
              }),
                console.log("==============" + a + "====================");
              console.log(error);
              console.log(names[a]);
              console.log("==============" + a + "====================");
              // flag = false;
              sendMail();
            } else {
              //Acknowledgement of sending mail with date and time
              console.log(
                `Email sent to => ${a + 1}/${names.length} ${
                  names[a]
                } at ${new Date().toLocaleString()}`
              );
              a++;
              b++;
              flag = true;
            }
          });
        }
      } else {
        mailOptions["to"] = "sanjaikumaran0311@gmail.com";
        subject = "Event announcement completed successfully!";

        mailOptions[
          "html"
        ] = `<div><h2>Hello SK,</h2><p>Mail automation for IEEE SB JEC about event annoucement is successfully completed.</p></div>`;
        clearInterval(interval); //To terminate the process after complete sending mail
      }
    }, 2000); //Giving 2 seconds of time interval to send each mail
  };
  sendMail("");

  var mailOptions = {
    from: "IEEE SB JEC<ieeesbjec@gmail.com>",
    to: "",
    cc: "",
    subject: subject,
    // attachments: [
    //   {
    //     filename: "", // Name of the attachment
    //     path: "", // Path to the file
    //   },
    // ],
    html: "", //content of the mail
  };
};
Mailer(
  ' Invitation to "Codemania: The Ultimate Code Battle" - Coding Competition by IEEE SB JEC'
);

const updateMails = (filePath) => {
  const fs = require("fs");
  var users = require("./mails.json");
  // Requiring the module
  const reader = require("xlsx");

  // Reading our test file
  const file = reader.readFile(filePath);

  let data = [];

  const sheets = file.SheetNames;

  for (let i = 0; i < sheets.length; i++) {
    const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
    temp.forEach((res) => {
      data.push(res);
    });
  }

  // Printing data

  // STEP 1: Reading JSON file
  // import users from "mails.json";
  //  data.forEach(  {
  //   for
  //   data=data.replace('First Name','Name')

  //  });
  // data.forEach((data) => {
  //   for (i = 0; i < 2; i++) {
  //     data[i] = data[i].replace("First Name", "Name");
  //     console.log(data[i]);
  //   }
  // });

  if (data[0]["First Name"]) {
    Name = "First Name";
  } else if (data[0]["Name"]) {
    Name = "Name";
  } else if (data[0]["name"]) {
    Name = "name";
  } else if (data[0]["first name"]) {
    Name = "first name";
  } else if (data[0]["First name"]) {
    Name = "First name";
  } else {
    Name = "NAMES";
  }
  if (data[0]["Email"]) {
    var Mail = "Email";
  } else if (data[0]["email"]) {
    Mail = "email";
  } else if (data[0]["Email Address"]) {
    Mail = "Email Address";
  } else if (data[0]["Email Address"]) {
    Mail = "Email Address";
  } else if (data[0]["email"]) {
    Mail = "email";
  } else if (data[0]["Mail"]) {
    Mail = "Mail";
  } else if (data[0]["mail"]) {
    Mail = "mail";
  } else if (data[0]["mail address"]) {
    Mail = "mail address";
  } else if (data[0]["Mail address"]) {
    Mail = "Mail address";
  } else if (data[0]["Mail address"]) {
    Mail = "Mail address";
  } else {
    Mail = "MAIL ID";
  }

  console.log(data[0]);
  // Defining new user
  // console.log(data);
  for (i = 0; i < data.length; i++) {
    var nm = data[i][Name].toLowerCase();
    var mail = data[i][Mail].toLowerCase();
    mail = mail.replace(" ", "");
    mail = mail.replace(",", "");
    mail = mail.replace("con", "com");

    nm = nm[0].toUpperCase() + nm.slice(1);
    if (users[nm] in users) {
      console.log("Already Present");
    } else {
      users[nm] = mail;
    }
  }
  // STEP 2: Adding new data to users object

  // STEP 3: Writing to a file
  fs.writeFile("mails.json", JSON.stringify(users), (err) => {
    // Checking for errors
    if (err) throw err;

    console.log("Done writing"); // Success
  });
};
// updateMails("C:/Users/DELL/Downloads/Students details.xlsx");
// updateMails("C:/Users/DELL/Downloads/a.xlsx");
// updateMails("C:/Users/DELL/Downloads/b.xlsx");
// updateMails("C:/Users/DELL/Downloads/c.xlsx");
// updateMails("C:/Users/DELL/Downloads/Book1.xlsx");
