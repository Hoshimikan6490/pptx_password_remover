const express = require("express");
const path = require("path");
const multer = require("multer");

const app = express();
const port = 8000;

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, "public/uploads/");
  },
  filename: function (req, file, cb) {
    cb(null, file.originalname);
  },
});

const upload = multer({ storage: storage });

app.use(express.static(path.join(__dirname, "public")));

app.get("/", (req, res) =>
  res.sendFile(path.join(__dirname, "public/pages/index.html"))
);

app.post("/", upload.single("file"), function (req, res) {
  res.send(req.file.originalname + "ファイルのアップロードが完了しました。\n処理を開始します…");
  
  
});

app.listen(port, function () {
  console.log(`Example app listening on port ${port}!`);
});
