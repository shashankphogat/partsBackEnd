require('dotenv').config();
const express =require("express");
const app = express();
const mongoose=require("mongoose");
const partsModel=require("./models/partsModel.js");
const partsRoutes=require("./routes/routes.js")
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
let allData;
const cors = require('cors');

const FRONTEND_URL = process.env.FRONTEND_URL || 'http://localhost:3000';
//middlewares
app.use(express.json());
app.use(express.urlencoded({extended:true}));
app.use(cors({
  origin: FRONTEND_URL,
  methods: ['GET','POST','PUT','DELETE'],
  credentials: true,
}));
//routes
app.use(partsRoutes)

//connect to database
mongoose.connect(process.env.MONGO_URI).then(()=>{
    app.listen(process.env.PORT, () => {
        console.log(`Server is listening at http://localhost:${process.env.PORT}`);
    });
}).catch((error)=>{
console.log(error)
});


// app.post("/api/formSubmitted",async(req,res)=>{
//     let {model,customer,havingDefect,typeOfDefect,sendPartFor} =req.body;
//     let parts= await partsModel.create(
//         {
//             model,
//             customer,
//             havingDefect,
//             typeOfDefect,
//             sendPartFor
//         }
//     )
//     allData=await partsModel.find()
//     res.send(allData);
// })


//create excel from mongodb data
// const ExcelJS = require('exceljs');

// const createExcelFromData = async (data) => {
//     const workbook = new ExcelJS.Workbook();
//     const worksheet = workbook.addWorksheet('Users');

//     // Add headers
//     worksheet.columns = [
//         { header: 'Model', key: 'model', width: 20 },
//         { header: 'Customer', key: 'customer', width: 30 },
//         { header: 'HavingDefect', key: 'havingDefect', width: 10 },
//         { header: 'TypeOfDefect', key: 'typeOfDefect', width: 20 },
//         { header: 'SendPartFor', key: 'sendPartFor', width: 20 }
//     ];

//     // Add rows
//     data.forEach(user => {
//         worksheet.addRow({
//             model: user.model,
//             customer: user.customer,
//             havingDefect: user.havingDefect,
//             typeOfDefect: user.typeOfDefect,
//             sendPartFor:user.sendPartFor
//         });
//     });

//     // Save the file
//     const filePath = './users.xlsx';
//     await workbook.xlsx.writeFile(filePath);
//     return filePath;
// };

// // nodemailer for sending mails
// const nodemailer = require('nodemailer');

// const transporter = nodemailer.createTransport({
//     service: 'gmail',
//     auth: {
//         user: 'shashankphogat26@gmail.com',
//         pass: 'ednf cczi tebu rtfu' // Use environment variables for security
//     }
// });

// const sendEmailWithAttachment = (to, subject, text, filePath) => {
//     const mailOptions = {
//         from: 'shashankphogat26@gmail.com',
//         to,
//         subject,
//         text,
//         attachments: [
//             {
//                 filename: 'users.xlsx',
//                 path: filePath
//             }
//         ]
//     };

//     transporter.sendMail(mailOptions, (error, info) => {
//         if (error) {
//             console.log('Error sending email:', error);
//         } else {
//             console.log('Email sent:', info.response);
//         }
//     });
// };


// // send mail on schedule basis with node cron
// const cron = require('node-cron');

// cron.schedule('48 20 * * *', async () => {
//     try {
//         // Fetch data from MongoDB
//         const users = await partsModel.find({});

//         // Create Excel file
//         const filePath = await createExcelFromData(users);

//         // Send email with the Excel file
//         sendEmailWithAttachment(
//             'nisar@nsidentics.com', // Replace with recipient email
//             'Daily User Report',
//             'Please find the attached user report.',
//             filePath
//         );

//         console.log('Scheduled task completed successfully.');
//     } catch (error) {
//         console.error('Error in scheduled task:', error);
//     }
// });


