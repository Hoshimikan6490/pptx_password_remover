const fs = require("fs");
const express = require("express");
const app = express();
const port = 8000;

app.get("/", (req, res) => {
  res.sendFile(__dirname + "/public/pages/index.html");
});

app.listen(port, () => {
  console.log(`listening on port ${port}`);
});
