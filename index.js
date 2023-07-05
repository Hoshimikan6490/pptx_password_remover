const express = require("express");
const path = require("path");
const fileUpload = require("express-fileupload");
const fs = require("fs");
const AdmZip = require("adm-zip");
const util = require("util");
const wait = util.promisify(setTimeout);

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

app.post("/", async function (req, res) {
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

  await wait(1000);
  fs.renameSync(
    `public/uploads/${file.name}`,
    `public/uploads/working.zip`,
    (err) => {
      if (err) throw err;
    }
  );

  const unzip = new AdmZip(path.join(__dirname, `public/uploads/working.zip`));
  unzip.extractAllTo(`public/uploads/editing`, true);

  let xmlString = fs.readFileSync(
    "public/uploads/editing/ppt/presentation.xml",
    "utf-8"
  );

  let newXmlString = remove_p_modifyVerifier(xmlString);

  fs.writeFileSync(`public/uploads/editing/ppt/presentation.xml`, newXmlString);

  fs.unlinkSync("public/uploads/working.zip");

  const zip = new AdmZip();
  // フォルダを追加
  zip.addLocalFolder("public/uploads/editing");

  // zipファイル書き出し
  zip.writeZip("public/uploads/finished.zip");

  fs.renameSync(
    `public/uploads/finished.zip`,
    `public/uploads/${file.name}`,
    (err) => {
      if (err) throw err;
    }
  );

  fs.rmSync("public/uploads/editing", { recursive: true, force: true });

  await wait(2000);
  res.download(
    path.join(__dirname, `public/uploads/${file.name}`),
    file.name,
    function (err) {
      if (err) {
        console.log(err);
      } else {
        fs.unlinkSync(`public/uploads/${file.name}`);
      }
    }
  );
});

app.listen(port, function () {
  console.log(`起動しました！　http://localhost:${port} をご覧ください!`);
});
