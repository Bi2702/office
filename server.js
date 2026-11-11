const express = require("express");
const https = require("https");
const path = require("path");
const devCerts = require("office-addin-dev-certs");

const app = express();
const port = Number(process.env.PORT || 3000);
app.use(express.static(path.join(__dirname, "public")));

devCerts
  .getHttpsServerOptions()
  .then((options) => {
    https.createServer(options, app).listen(port, () => {
      console.log(`Add-in đang chạy tại https://localhost:${port}`);
    });
  })
  .catch((error) => {
    console.error("Không tạo được chứng chỉ HTTPS cho máy phát triển.", error);
    process.exitCode = 1;
  });
