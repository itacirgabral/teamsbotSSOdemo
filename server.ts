import * as Express from "express";
import * as http from "http";
import * as path from "path";
import { MsTeamsApiRouter, MsTeamsPageRouter } from "express-msteams-host";
import * as debug from "debug";

const log = debug("msteams");

log("Initializing Microsoft Teams Express hosted App...");

require("dotenv").config();

import * as allComponents from "./TeamsAppsComponents";

const express = Express();
const port = process.env.port || process.env.PORT || 3007;

express.use(Express.json({
  verify: (req, res, buf: Buffer, encoding: string): void => {
    (req as any).rawBody = buf.toString();
  }
}));
express.use(Express.urlencoded({ extended: true }));

// spy
express.use(function (req, res, next) {
  console.log("#######")
  console.log("# SPY #")
  console.log("#######")
  console.log(JSON.stringify(req.body, null, 2))
  next()
})

express.use(MsTeamsApiRouter(allComponents));

express.use(MsTeamsPageRouter({
  root: path.join(__dirname, "web/"),
  components: allComponents
}));

express.set("port", port);

http.createServer(express).listen(port, () => {
  log(`Server running on ${port}`);
});
