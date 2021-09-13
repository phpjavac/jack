require("dotenv").config();
const creds = require("./config/jack-325902-b64e06004008.json");
const schedule = require("node-schedule");
const { GoogleSpreadsheet } = require("google-spreadsheet");
const request = require("request");
const jsdom = require("jsdom");


const docId = process.env.DOCID;
const user = require("./user");
const getQueryVariable = (variable, url) => {
  var query = url.substring(1);
  var vars = query.split("&");
  for (var i = 0; i < vars.length; i++) {
    var pair = vars[i].split("=");
    if (pair[0] == variable) {
      return pair[1];
    }
  }
  return false;
};
const jira = (name, task, task_h, row) => {
  const monthList = [
    "一月",
    "二月",
    "三月",
    "四月",
    "五月",
    "六月",
    "七月",
    "八月",
    "九月",
    "十月",
    "十一月",
    "十二月",
  ];
  const t = new Date().getTime();
  // 86400000 一天的时间戳 * 2 计划完成时间推迟两天
  const date = new Date(t + 86400000 * 2);

  const year = date.getFullYear().toString().slice(2);
  const month = monthList[date.getMonth()];
  const day = date.getDate();
  const time = `${day}/${month}/${year}`;

  const headers = {
    "User-Agent":
      "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.181 Safari/537.36",
    Accept:
      "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
    cookie: [
      process.env.JSESSIONID,
      process.env.token
    ],
  };
  const url = `${process.env.JIRA_URL}browse/${task}`;
  request(
    {
      url,
      method: "GET",
      headers,
    },
    (err, response, body) => {
      console.log(`开始分配${name}的任务一${task}`);
      const dom = new jsdom.JSDOM(body);

      const type = dom.window.document.getElementById("type-val");
      if (!type) {
        console.error(`分配失败--检查任务状态-type`);
        console.error(`分配${name}的任务一${task}失败。`);
        row.状态 = "分配失败";
        row.save();
        return;
      }
      if (!dom.window.document.getElementById("action_id_711")) {
        console.error(`分配失败--检查任务状态`);
        console.error(`分配${name}的任务一${task}失败。`);
        row.状态 = "分配失败";
        row.save();
        return;
      }
      const typeInfo = type.innerHTML;
      let toURl = dom.window.document.getElementById("action_id_711").href;
      toURl = toURl.slice(toURl.indexOf("?"));
      const token = getQueryVariable("atl_token", toURl);
      const id = getQueryVariable("id", toURl);
      const url = `${
        process.env.JIRA_URL
      }/secure/CommentAssignIssue.jspa?atl_token=${encodeURI(token)}`;
      const form = {
        inline: true,
        decorator: "dialog",
        action: 711,
        id: id,
        viewIssueKey: "",
        customfield_10905: time,
        customfield_10100: user[name],
        customfield_10101: task_h,
        comment: "",
        commentLevel: "",
        atl_token: token,
      };
      if (typeInfo.includes("缺陷")) {
        form.customfield_10600 = 10102;
      }

      request(
        {
          url,
          method: "POST",
          headers,
          form,
        },
        (err) => {
          if (err) {
            console.error(`分配${name}的任务一${name}失败。`);
            row.状态 = "分配失败";
            row.save();
            return;
          }
          row.状态 = "分配成功";
          row.save();
        }
      );
    }
  );
};
const init = async () => {
  const doc = new GoogleSpreadsheet(docId);
  await doc.useServiceAccountAuth(creds);
  await doc.loadInfo();
  for (let index = 0; index < doc.sheetsByIndex.length; index++) {
    const sheet = doc.sheetsByIndex[index];
    const rows = await sheet.getRows();
    for (let idx = 0; idx < rows.length; idx++) {
      const row = rows[idx];
      if(row.状态 !== "分配失败" || row.状态 !== "分配成功"){
          jira(sheet.title, row.任务号, row.计划开发工时, row);
      }
    }
  }
};

init();
// 启动任务
let job = schedule.scheduleJob("*/5 * * * *", () => {
  console.log("开始分配");
  init();
});
