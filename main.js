const express = require('express');
const axios = require('axios');
const FormData = require('form-data');
const fs = require('fs');
const path = require('path');
const { saveAssociationToFile, getEmailIDFromFile} = require('./keyValueStore');
const {emailService} = require('./emailService')

const app = express();
const port = 3000;



app.use(express.json());


app.post('/sendPdf', async (req, res) => {
  try {
    const { headers } = req;
    const { docid, filepath, url1, url2 } = headers;
    
    console.log('Received Headers:');
    console.log('docId:', docid);
    console.log('filePath:', filepath);
    console.log('url1:', url1);
    console.log('url2:', url2);

    const cleanedFilePath = filepath.replace(/'/g, '');

    //Email function
    const agentEmailID = getEmailIDFromFile(docid);
    await emailService(agentEmailID,'FORM-1770S_'+docid,cleanedFilePath);

    //

    await makeRequest(url1, docid, cleanedFilePath);
    await makeRequest(url2, docid, cleanedFilePath);

    res.status(200).json({ success: true });
  } catch (error) {
    console.error('Error:', error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

async function makeRequest(url, docid, filepath) {
  const formData = new FormData();
  formData.append('docId', docid);
  formData.append('file', fs.createReadStream(filepath));

  const config = {
    method: 'post',
    url,
    headers: formData.getHeaders(),
    data: formData,
    maxContentLength: Infinity,
    maxBodyLength: Infinity,
  };

  const response = await axios(config);
  console.log(`Response from ${url}:`, response.data);
}

// app.post('/sendPdf', async (req, res) => {
//   try {
//     const { headers } = req;
//     const { docid, filepath, url1, url2 } = headers;



//     // Log the received headers
//     console.log('Received Headers:');
//     console.log('docId:', docid);
//     console.log('filePath:', filepath);
//     console.log('url1:', url1);
//     console.log('url2:', url2);

//     // Remove any single quotes from the file path
//     const cleanedFilePath = filepath.replace(/'/g, '');

//     const agentEmailID = getEmailIDFromFile(docid)

//     await emailService(agentEmailID,'FORM-1770S_'+docid,cleanedFilePath)

//     // Log the cleaned file path
//     console.log('Cleaned FilePath:', cleanedFilePath);

//     // Create FormData object for the first URL
//     const data1 = new FormData();
//     data1.append('docId', docid);
//     data1.append('file', fs.createReadStream(cleanedFilePath));

//     // Axios configuration for the first URL
//     const config1 = {
//       method: 'post',
//       url: url1,
//       headers: {
//         ...data1.getHeaders(),
//       },
//       data: data1,
//       maxContentLength: Infinity,
//       maxBodyLength: Infinity,
//     };

//     // Make Axios request for the first URL
//     const response1 = await axios(config1);

//     console.log('Response from URL 1:', response1.data);

//     // Create FormData object for the second URL
//     const data2 = new FormData();
//     data2.append('docId', docid);
//     data2.append('file', fs.createReadStream(cleanedFilePath));

//     // Axios configuration for the second URL
//     const config2 = {
//       method: 'post',
//       url: url2,
//       headers: {
//         ...data2.getHeaders(),
//       },
//       data: data2,
//       maxContentLength: Infinity,
//       maxBodyLength: Infinity,
//     };

//     // Make Axios request for the second URL
//     const response2 = await axios(config2);

//     console.log('Response from URL 2:', response2.data);

//     res.status(200).json({ success: true });
//   } catch (error) {
//     console.error('Error:', error.message);
//     res.status(500).json({ success: false, error: error.message });
//   }
// });


app.post('/form1770s', (req, res) => {
  const { headers } = req;
  const ticketId = headers['ticketid'];
  const fileExtension = headers['fileextension'];
  const userEmail = headers['useremail']
  
// save email and ticketID
 saveAssociationToFile(ticketId, userEmail);

  console.log('Main 1770s Method');

  let body = [];

  req.on('data', (chunk) => {
    body.push(chunk);
  });

  req.on('end', () => {
    body = Buffer.concat(body);

    const result = 'File has been received successfully';

    createActualFile(body, ticketId, fileExtension);

    res.type('plain/text').send(result);
  });
});

function createActualFile(body, ticketId, fileExtension) {
  const folderPath = `C:\\FORM1770S\\Unprocessed\\${ticketId}`;
  const filePath = path.join(folderPath, `${ticketId}.${fileExtension}`);

  if (!fs.existsSync(folderPath)) {
    fs.mkdirSync(folderPath, { recursive: true });
  }

  fs.writeFileSync(filePath, body);
  console.log(`File saved at ${filePath}`);
}


app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
  console.log(`Exposed API Endpoints:`);
  console.log(`1. PDF Upload Endpoint: http://localhost:${port}/sendPdf`);
  console.log(`2. Form Data Upload Endpoint: http://localhost:${port}/form1770s`);
});

