# Node XLSX

此 api 是通过 node-xlsx 、xlsx-style、xlsx 三个 npm 组合使用的，来弥补 node-xlsx 不能设置颜色的问题，参数格式按照 node-xlsx 方式使用，设置样式参数按照 xlsx-style。

注意：
存在缺陷 设置不了行高(待完善)

## 按照

`npm i xlsx-style-node`

## 使用

在 js 中使用
`const xlsx = require("xlsx-style-node").default;`

在 ts 中使用
`import xlsxStyleNode from 'xlsx-style-node';`

```js
const fs = require("fs");
const data: any = [
  ["导表时间：2021-05-27 15:36:28"],
  ["用户Id: 288884"],
  ["用户名：test"],
  [],
  ["对账公式："],
  ["期末余额=期初余额+入金-出金"],
  [],
  ["币种：USD"],
  [
    "*期初余额指本月度1号0点0分0秒余额",
    null,
    null,
    "*期末余额指本月度最后一天23点59分59秒余额",
  ],
  [
    "期初余额（总计）",
    "10000.01",
    null,
    "期末余额（总计）",
    "20000.01",
    null,
    "入金（总计）",
    "20000",
    null,
    "出金（总计）",
    "10000",
  ],
  [
    "账户余额",
    "8000",
    null,
    "账户余额",
    "18000",
    null,
    "账户充值",
    "19000",
    null,
    "账户转出",
    "9000",
  ],
  [
    "储值卡总可用余额",
    "99",
    null,
    "储值卡总可用余额",
    "3232",
    null,
    "账户转入",
    "999",
    null,
    "账户充值手续费",
    "000",
  ],
  [
    "额度卡（预算）总可用余额",
    "99",
    null,
    "额度卡（预算）总可用余额",
    "33",
    null,
    "储值卡转入",
    "44",
    null,
    "储值卡转出",
    "500",
  ],
  [
    null,
    null,
    null,
    null,
    null,
    null,
    "储值卡退款",
    "22",
    null,
    "储值卡消费（已完成）",
    "111",
  ],
  [
    null,
    null,
    null,
    null,
    null,
    null,
    "额度卡转入",
    "22",
    null,
    "储值卡消费（处理中）",
    "111",
  ],
  [
    null,
    null,
    null,
    null,
    null,
    null,
    "额度卡退款",
    "22",
    null,
    "储值卡退款手续费",
    "111",
  ],
  [null, null, null, null, null, null, null, null, null, "额度卡转出", "11"],
  [
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    "额度卡消费（已完成）",
    "11",
  ],
  [
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    "额度卡消费（处理中）",
    "11",
  ],
  [
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    "额度卡退款手续费",
    "11",
  ],
  [null, null, null, null, null, null, null, null, null, "开卡费", "11"],
];

// 设置样式
for (const index in data) {
  if (Number(index) < 5) continue;
  if (index === "7") continue; // 第八行不处理
  if (!data[index].length) continue;
  for (const index2 in data[index]) {
    if (!data[index][index2]) continue;
    const content = data[index][index2];
    if (index === "5") {
      data[index][index2] = {
        v: content,
        s: {
          font: {
            size: 24,
            bold: true, // 加粗
          },
        },
      };
    } else {
      data[index][index2] = {
        v: content,
        s: {
          alignment: {
            vertical: "center",
            horizontal: "center",
          },
          font: {
            size: 24,
            ...(index === "8" && { color: { rgb: "ff280c" } }), // 设置颜色
            ...(index === "9" && { bold: true }), // 字体加粗
          },
        },
      };
    }
  }
}

// 单元格宽度
const options = {
  "!cols": [
    { wpx: 260 },
    { wpx: 80 },
    { wpx: 30 },
    { wpx: 300 },
    { wpx: 80 },
    { wpx: 30 },
    { wpx: 160 },
    { wpx: 80 },
    { wpx: 30 },
    { wpx: 170 },
    { wpx: 80 },
  ],
};

const buffer = xlsx.build([{ name: "mySheetName", data: data, options }]);
fs.writeFileSync("test1.xlsx", buffer);
```

## 文档

[node-xlsx](https://www.npmjs.com/package/node-xlsx)
