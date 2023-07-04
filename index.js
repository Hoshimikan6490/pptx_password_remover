const express = require("express");
const path = require("path");
const multer = require("multer");
const fs = require("fs");
const AdmZip = require("adm-zip");

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

function remove_p_modifyVerifier(xml) {
  let i = 0;
  while (1) {
    i++;
    if (xml.substr(i, 1) == "p") {
      if (xml.substr(i + 2, 14) == "modifyVerifier") {
        var first_hand = xml.substr(0, i - 1);

        for (j = i; i <= 6000; j++) {
          if (xml.substr(j, 1) == ">") {
            var latter_hand = xml.slice(j + 1);
            break;
          }
        }

        return first_hand + latter_hand;
      }
    }
  }
}

app.post("/", upload.single("file"), function (req, res) {
  res.send(
    req.file.originalname +
      "ファイルのアップロードが完了しました。\n処理を開始します…"
  );

  fs.renameSync(
    `public/uploads/${req.file.originalname}`,
    `public/uploads/working.zip`,
    (err) => {
      if (err) throw err;
    }
  );

  const zip = new AdmZip(path.join(__dirname, `public/uploads/working.zip`));
  zip.extractAllTo(`public/uploads/editing`, true);

  let xmlString = fs.readFileSync(
    "public/uploads/editing/ppt/presentation.xml",
    "utf-8"
  );

  let newXmlString = remove_p_modifyVerifier(xmlString);

  fs.writeFileSync(`public/uploads/editing/ppt/presentation.xml`, newXmlString);

  fs.unlinkSync("public/uploads/working.zip");
});

app.listen(port, function () {
  console.log(`Example app listening on port ${port}!`);
});
