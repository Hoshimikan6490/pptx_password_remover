const express = require("express");
const path = require("path");
const fileUpload = require("express-fileupload");
const fs = require("fs");
const AdmZip = require("adm-zip");

const app = express();
const port = 8000;

app.use(
  fileUpload({
    defCharset: "utf8",
    defParamCharset: "utf8",
  })
);

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

app.post("/", function (req, res) {
  if (!req.files || Object.keys(req.files).length === 0) {
    res.status(400).send("No files were uploaded.");
    return;
  }

  let file = req.files.file;
  let uploadPath = __dirname + "/public/uploads/" + file.name;

  file.mv(uploadPath, function (err) {
    if (err) {
      return res.status(500).send(err);
    }
  });

  res.send(
    file.name +
      "ファイルのアップロードが完了しました。\n処理を開始します…"
  );

  fs.renameSync(
    `public/uploads/${file.name}`,
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
  console.log(`The app listening on port ${port}!`);
});
