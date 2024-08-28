const AWS = require("aws-sdk");
const JSZip = require("jszip");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

// Configure AWS SDK with your credentials
AWS.config.update({
  accessKeyId: "AWS_ACCESS_KEY_ID",
  secretAccessKey: "SECRET_ACCESS_KEY",
  region: "AWS_REGION",
});

const s3 = new AWS.S3();

const bucketName = "BUCKETNAME";
const uploadDirectory = "SiteImages"; // Optional, directory in S3 to upload files to
let listOfImgToUpload = [];

// Check if file is an Excel file
function isExcelFile(filename) {
  const ext = path.extname(filename).toLowerCase();
  return ext === ".xlsx" || ext === ".xls" || ext === ".xlsm";
}

// Read image names from the Excel file buffer
async function readImageNameColumn(buffer) {
  try {
    const workbook = xlsx.read(buffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    const headers = jsonData[0];
    const imageNameIndex = headers.indexOf("Image Name");

    if (imageNameIndex === -1) {
      console.log('Column "Image Name" not found.');
      return;
    }

    const imageNames = jsonData
      .map((row) => row[imageNameIndex])
      .filter((_, index) => index > 0);

    listOfImgToUpload = imageNames.filter((img) => img !== undefined);
  } catch (error) {
    console.error("Error reading image names from Excel file:", error);
  }
}

// Process ZIP file
async function processZipFile() {
  try {
    if (!fs.existsSync("./sites.zip")) {
      throw new Error("ZIP file not found.");
    }

    const data = fs.readFileSync("./sites.zip");
    const zip = new JSZip();
    const contents = await zip.loadAsync(data);

    // Extract image names from Excel files
    await Promise.all(
      Object.keys(contents.files)
        .filter((filename) => isExcelFile(filename))
        .map((filename) =>
          zip.file(filename).async("nodebuffer").then(readImageNameColumn)
        )
    );

    // Log the list of images to upload
    console.log(listOfImgToUpload);

    // Iterate through other files in the ZIP
    await Promise.all(
      Object.keys(contents.files).map(async (filename) => {
        if (!contents.files[filename].dir) {
          const content = await contents.files[filename].async("nodebuffer");
          if (listOfImgToUpload.includes(filename.split(".")[0])) {
            console.log(filename, "files to upload");

            // Upload the file to S3
            const params = {
              Bucket: bucketName,
              Key: path.join(uploadDirectory, filename),
              Body: content,
            };
            try {
              const data = await s3.upload(params).promise();
              console.log(`File uploaded successfully at ${data.Location}`);
            } catch (s3Err) {
              console.error(`Error uploading file ${filename}:`, s3Err);
            }
          }
        }
      })
    );
  } catch (err) {
    console.error("Error processing ZIP file:", err);
  }
}

processZipFile().catch((err) => console.error("Unhandled error:", err));

