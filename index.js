const express = require("express");
const app = express();

const request = require("superagent");
const cheerio = require("cheerio");
require("superagent-charset")(request);
const urlencode = require("urlencode-gb2312-ignore");
const nodeExcel = require("excel-export");

let server = app.listen(4000, function () {
    let host = server.address().address;
    let port = server.address().port;
    console.log("Your App is running at http://%s:%s", host, port);
});

let alist = [
    "北京市（京）",
    "天津市（津）",
    "河北省（冀）",
    "山西省（晋）",
    "内蒙古自治区（内蒙古）",
    "辽宁省（辽）",
    "吉林省（吉）",
    "黑龙江省（黑）",
    "上海市（沪）",
    "江苏省（苏）",
    "浙江省（浙）",
    "安徽省（皖）",
    "福建省（闽）",
    "江西省（赣）",
    "山东省（鲁）",
    "河南省（豫）",
    "湖北省（鄂）",
    "湖南省（湘）",
    "广东省（粤）",
    "广西壮族自治区（桂）",
    "海南省（琼）",
    "重庆市（渝）",
    "四川省（川、蜀）",
    "贵州省（黔、贵）",
    "云南省（滇、云）",
    "西藏自治区（藏）",
    "陕西省（陕、秦）",
    "甘肃省（甘、陇）",
    "青海省（青）",
    "宁夏回族自治区（宁）",
    "新疆维吾尔自治区（新）",
    "香港特别行政区（港）",
    "澳门特别行政区（澳）",
    "台湾省（台）",
];

let list = [];

let requestFn = (name) => {
    return new Promise((resolve) => {
        request
            .get("http://xzqh.mca.gov.cn/defaultQuery?shengji=" + urlencode(name, "gb2312") + "&diji=-1&xianji=-1")
            .charset("gbk")
            .end((err, res) => {
                if (err) {
                    console.log(err);
                } else {
                    list = getData(res, name);
                    resolve(list);
                }
            });
    });
};

let getData = (res, name) => {
    let data = [];
    let $ = cheerio.load(res.text, { decodeEntities: false });
    $(".shi_nub").each((idx, ele) => {
        let item = {};
        $("td", ele).each((i, e) => {
            switch (i) {
                case 0:
                    item.diming = $(e).text();
                    break;
                case 1:
                    item.zhudi = $(e).text();
                    break;
                case 2:
                    item.renkou = $(e).text();
                    break;
                case 3:
                    item.mianji = $(e).text();
                    break;
                case 4:
                    item.daima = $(e).text();
                    break;
                case 5:
                    item.quhao = $(e).text();
                    break;
                case 6:
                    item.youbian = $(e).text();
                    break;
                default:
                    break;
            }
        });
        let child = [];
        $("tr[parent=" + $(ele).attr("flag") + "][type=2]").each((iii, eee) => {
            let childitem = {};
            $("td", eee).each((ii, ee) => {
                switch (ii) {
                    case 0:
                        childitem.diming = $(ee).text();
                        break;
                    case 1:
                        childitem.zhudi = $(ee).text();
                        break;
                    case 2:
                        childitem.renkou = $(ee).text();
                        break;
                    case 3:
                        childitem.mianji = $(ee).text();
                        break;
                    case 4:
                        childitem.daima = $(ee).text();
                        break;
                    case 5:
                        childitem.quhao = $(ee).text();
                        break;
                    case 6:
                        childitem.youbian = $(ee).text();
                        break;
                    default:
                        break;
                }
            });
            child.push(childitem);
        });
        data.push({ parent: item, child: child });
    });
    console.log(name + "--->完成");
    return data;
};

app.get("/excel/:id", async (req, res, next) => {
    await requestFn(alist[req.params.id]);
    console.log(req.params.id);
    let name = encodeURI("全国行政区表" + alist[req.params.id]);
    let conf = {};
    let row = [];
    conf.name = "first";
    const colsArr = [
        { caption: "地名", type: "string" },
        { caption: "驻地", type: "string" },
        { caption: "人口（万人）", type: "string" },
        { caption: "面积（平方千米）", type: "string" },
        { caption: "行政区划代码", type: "string" },
        { caption: "区号", type: "string" },
        { caption: "邮编", type: "string" },
    ];
    conf.cols = colsArr;
    list.forEach((item, index) => {
        let tmp = [];
        tmp.push(item.parent.diming);
        tmp.push(item.parent.zhudi);
        tmp.push(item.parent.renkou);
        tmp.push(item.parent.mianji);
        tmp.push(item.parent.daima);
        tmp.push(item.parent.quhao);
        tmp.push(item.parent.youbian);
        row.push(tmp);
        item.child.forEach((sitem, sindex) => {
            let stmp = [];
            stmp.push(sitem.diming);
            stmp.push(sitem.zhudi);
            stmp.push(sitem.renkou);
            stmp.push(sitem.mianji);
            stmp.push(sitem.daima);
            stmp.push(sitem.quhao);
            stmp.push(sitem.youbian);
            row.push(stmp);
        });
    });
    conf.rows = row;
    let result = nodeExcel.execute(conf);

    res.setHeader("Content-Type", "application/vnd.openxmlformats;charset=utf-8");
    res.setHeader("Content-Disposition", "attachment; filename=" + name + ".xlsx");
    res.end(result, "binary");
});
