require('dotenv').config();
const express = require("express");
const app = express();
const bodyParser = require("body-parser");
const cors = require("cors");
const request = require("request");

app.use(cors());
app.use(bodyParser.json());

const realendpoint = "https://api.clashofclans.com/v1/players/";

app.get('/:cleanTag', async (req, res) => {
  let tag = req.params.cleanTag;
  if (!tag.startsWith('#')) {
    tag = '#' + tag;
  }
  const url = realendpoint + encodeURIComponent(tag);
  console.log(`Fetching: ${url}`);
  
  request.get({
    url: url,
    headers: {'Authorization': `Bearer ${process.env.CLASH_API_KEY}`}
  }, function (err, resp, body) {
    if(err) {
      console.error('Request error:', err);
      res.status(500).json({error: err.message});
    } else {
      console.log(`Response status: ${resp.statusCode}`);
      console.log('Response body:', body);
      try {
        const data = JSON.parse(body);
        res.status(resp.statusCode).json(data);
      } catch (parseErr) {
        console.error('Parse error:', parseErr);
        res.status(resp.statusCode).send(body);
      }
    }
  });
});

app.listen(4000, () =>
  console.log(`Wrapper API Server Running On Port -> 4000!`)
);