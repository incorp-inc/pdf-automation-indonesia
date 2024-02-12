const nodemailer = require('nodemailer');
require('dotenv').config();

// Create a transporter object using SMTP transport
let transporter = nodemailer.createTransport({
    // host: 'smtp.office365.com',
    // port: 587,
    // secure: false,
    // auth: {
    //   user: 'notifications@incorp.asia',
    //   pass: 'Rikvin1234',
    // },

    host: process.env.MAIL_HOST,
    port: process.env.MAIL_PORT,
    secure: false,
    auth: {
      user: process.env.MAIL_USERNAME,
      pass: process.env.MAIL_PASSWORD,
    },
  });


module.exports.emailService = async(email,fileName,filePath) => {

    console.log('Path',filePath)

let mailOptions = {
    from: 'notifications@incorp.asia',
    to: email,
    subject: 'Email with Form-1770S Attachment ',
    text: 'Please find the attached file.',
    attachments: [
        {
            filename: fileName+'.pdf',
            path: filePath 
        }
    ]
};

// Send email
transporter.sendMail(mailOptions, (error, info) => {
    if (error) {
        console.log('Error occurred:', error);
    } else {
        console.log('Email sent:', info.response);
    }
});

}
