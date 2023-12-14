const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const xlsx = require('xlsx');
const PDFDocument = require('pdfkit');
const nodemailer = require('nodemailer');
const fs = require('fs');
const app = express();
const PORT = 3000;

require('dotenv').config();

app.set('view engine', 'ejs');
app.set('views', __dirname + '/views');

const DATABASE_NAME = 'new';
const MONGODB_URI = `mongodb+srv://Harshit:Harshit@mental.wlgbwhj.mongodb.net/${DATABASE_NAME}?retryWrites=true&w=majority`;

mongoose.connect(MONGODB_URI, {
    useNewUrlParser: true,
    useUnifiedTopology: true,
});

const db = mongoose.connection;

db.on('error', (err) => {
    console.error(`MongoDB connection error: ${err}`);
});

db.once('open', () => {
    console.log(`Connected to MongoDB: ${DATABASE_NAME}`);
});

const Schema = mongoose.Schema;

const excelSchema = new Schema({
    name: String,
    email: String,
    mobile: String,
    amount: Number,
    numberoftree: Number,
});

const ExcelModel = mongoose.model('ExcelModel', excelSchema);

app.use(express.json());
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });


app.get('/', (req, res) => {
    res.render('upload', { message: null });
});


app.post('/upload', upload.single('excelFile'), async (req, res) => {
    try {
        const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet);

        const certificates = data.map((row) => ({
            name: row.name,
            email: row.email,
            mobile: row.mobile,
            amount: row.amount,
            numberoftree: row.numberoftree,
        }));

        await ExcelModel.create(certificates);

     
        res.redirect('/email-page');
        console.log("Uploaded");
    } catch (error) {
        console.error(error);
        res.status(500).send('Internal Server Error');
    }
});


app.get('/email-page', (req, res) => {
    res.render('email-page');
});

app.post('/send-emails', async (req, res) => {
    try {
       
        const mongoData = await ExcelModel.find();

      
        const transporter = nodemailer.createTransport({
            service: 'gmail',
            auth: {
                user: process.env.EMAIL_USER,
                pass: process.env.EMAIL_PASSWORD,
            },
        });
        
        for (const userData of mongoData) {
            
            const pdfDoc = new PDFDocument({ margin: 10, size: [1273, 771] });

          


          const backgroundImagePath = './public/custom.jpg'; 
          pdfDoc.image(backgroundImagePath, 0, 0, { width: 1273, height: 771 });

          pdfDoc.font('./public/segoesc.ttf').fontSize(50);
          pdfDoc.text(`${userData.name}`, 600, 300); 
          pdfDoc.font('./public/segoesc.ttf').fontSize(20);
                      pdfDoc.text(` ${userData.numberoftree}`, 658, 394); 

            pdfDoc.fontSize(12).font('Helvetica');



            const pdfFilePath = `./public/certificates_${userData.name}.pdf`; 
            pdfDoc.pipe(fs.createWriteStream(pdfFilePath));
            pdfDoc.end();

            
            const pdfBuffer = await new Promise((resolve) => {
                const chunks = [];
                pdfDoc.on('data', (chunk) => chunks.push(chunk));
                pdfDoc.on('end', () => resolve(Buffer.concat(chunks)));
            });

           
            const mailOptions = {
                from: process.env.EMAIL_USER,
                to: userData.email,
                subject: 'Tree Planting Certificate',
                text: `Dear ${userData.name},\n\nAttached is your Tree Planting Certificate.\n\nRegards,\nRockstar Games`,
                attachments: [
                    {
                        filename: 'Tree_Planting_Certificate.pdf',
                        content: pdfBuffer, 
                    },
                ],
            };

          
            await transporter.sendMail(mailOptions);
        }

        res.send('Emails sent successfully!');
    } catch (error) {
        console.error(error);
        res.status(500).send('Internal Server Error');
    }
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
