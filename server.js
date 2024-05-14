import express from 'express'
import multer from 'multer';
import xlsx from 'xlsx';
import {MongoClient} from 'mongodb';
import mongoose from 'mongoose';
import cors from 'cors';
import { ObjectId } from 'mongodb';

const app = express();
const upload = multer({ dest: 'uploads/' });

const url = 'mongodb+srv://athrihegde:athrihegde@cluster0.7agvvhy.mongodb.net/?retryWrites=true&w=majority';

app.use(express.json());
const corsOptions = {
    origin: '*', // specific origin
    credentials: true // allow credentials (cookies, authorization headers, etc.)
  };
  
  app.use(cors(corsOptions));
  
mongoose.connect(url, {
  useNewUrlParser: true,
  
});


mongoose.connect(url, {
  useNewUrlParser: true,
  useUnifiedTopology: true,
});
const conn = mongoose.connection;
conn.on('error', console.error.bind(console, 'MongoDB connection error:'));
conn.once('open', () => {
  console.log('Connected to MongoDB');
});



app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('No file uploaded');
    }

    const workbook = xlsx.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const excelData = xlsx.utils.sheet_to_json(sheet);

    const client = new MongoClient(url, { useUnifiedTopology: true });
    await client.connect();

    const dbName = 'excel_db';
    const db = client.db(dbName);
    const collectionName = 'excel_collection';
    const collection = db.collection(collectionName);

    const result = await collection.insertMany(excelData);
    console.log(`${result.insertedCount} records inserted into MongoDB`);

    const insertedData = await collection.find({}).toArray();

    res.json(insertedData);

    await client.close();
  } catch (error) {
    console.error('Error:', error);
    res.status(500).send('Internal Server Error');
  } finally {
    if (req.file) {
      const fs = require('fs');
      fs.unlinkSync(req.file.path);
    }
  }
});

app.get('/download', async (req, res) => {
  try {
    const client = new MongoClient(url, { useUnifiedTopology: true });
    await client.connect();

    const dbName = 'excel_db';
    const db = client.db(dbName);

    const collectionName = 'excel_collection';
    const collection = db.collection(collectionName);

    const data = await collection.find({}).toArray();

    const worksheet = [
      ['Roll No', 'Name', 'Sem'],
      ...data.map(item => [item['Roll No'], item['Name'], item['Sem']])
    ];

    const sheet = xlsx.utils.aoa_to_sheet(worksheet);
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, sheet, 'Sheet1');

    const excelBuffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
    res.set('Content-Disposition', 'attachment; filename="data.xlsx"');
    res.set('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(excelBuffer);

    await client.close();
  } catch (error) {
    console.error('Error:', error);
    res.status(500).send('Internal Server Error');
  }
});

app.get('/download-format', async (req, res) => {
  try {
    const worksheet = [
      ['Roll No', 'Name', 'Sem'],
    ];

    const sheet = xlsx.utils.aoa_to_sheet(worksheet);
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, sheet, 'Sheet1');

    const excelBuffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
    res.set('Content-Disposition', 'attachment; filename="excel_format.xlsx"');
    res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(excelBuffer);
  } catch (error) {
    console.error('Error:', error);
    res.status(500).send('Internal Server Error');
  }
});

app.delete('/delete/:id', async (req, res) => {
  try {
    const client = new MongoClient(url, { useUnifiedTopology: true });
    await client.connect();

    const dbName = 'excel_db';
    const db = client.db(dbName);

    const collectionName = 'excel_collection';
    const collection = db.collection(collectionName);

    const id = req.params.id;
    const query = { _id: new ObjectId(id) }; // Use new ObjectId to create an instance of ObjectId

    await collection.deleteOne(query);

    const data = await collection.find({}).toArray();
    res.json(data);

    await client.close();
  } catch (error) {
    console.error('Error:', error);
    res.status(500).send('Internal Server Error');
  }
});

app.get('/fetch-data', async (req, res) => {
  try {
    const client = new MongoClient(url, { useUnifiedTopology: true });
    await client.connect();

    const dbName = 'excel_db';
    const db = client.db(dbName);

    const collectionName = 'excel_collection';
    const collection = db.collection(collectionName);

    const data = await collection.find({}).toArray();

    res.json(data);

    await client.close();
  } catch (error) {
    console.error('Error:', error);
    res.status(500).send('Internal Server Error');
  }
});


app.post('/save-edit', async (req, res) => {
  try {
    const client = new MongoClient(url, { useUnifiedTopology: true });
    await client.connect();

    const dbName = 'excel_db';
    const db = client.db(dbName);

    const collectionName = 'excel_collection';
    const collection = db.collection(collectionName);

    const editedData = req.body.editedData.map(row => ({
      ...row,
      _id: new ObjectId(row._id), // Use new ObjectId to create an instance of ObjectId
    }));

    await Promise.all(editedData.map(async row => {
      const filter = { _id: row._id };
      const update = { $set: row };
      await collection.updateOne(filter, update);
    }));

    const data = await collection.find({}).toArray();
    res.json(data);

    await client.close();
  } catch (error) {
    console.error('Error:', error);
    res.status(500).send('Internal Server Error');
  }
});
