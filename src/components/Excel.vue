<template>
  <div class="box">
    <button type="button" @click="tableToExcel">导出1</button>
    <button type="button" @click="tableToExcels">导出2</button>
    <button type="button" @click="exportExcle">导出3</button>
    <button type="button" @click="exportcc">导出4</button>
  </div>
</template>

<script setup>
import { ref } from "vue";
import * as XLSX from "xlsx";
import XLSXS from "xlsx-js-style";

const tableToExcel = () => {
  // 要导出的json数据
  const jsonData = [
    {
      name: "路人甲",
      phone: "123456",
      email: "123@123456.com",
    },
    {
      name: "炮灰乙",
      phone: "123456",
      email: "123@123456.com",
    },
    {
      name: "土匪丙",
      phone: "123456",
      email: "123@123456.com",
    },
    {
      name: "流氓丁",
      phone: "123456",
      email: "123@123456.com",
    },
  ];
  // 列标题
  let str = `<tr><th style="color:red;">姓名</th><th>电话</th><th>邮箱</th></tr>`;
  // 循环遍历，每行加入tr标签，每个单元格加td标签
  for (let i = 0; i < jsonData.length; i++) {
    str += "<tr>";
    for (const key in jsonData[i]) {
      // 增加\t为了不让表格显示科学计数法或者其他格式
      str += `<td>${jsonData[i][key] + "\t"}</td>`;
    }
    str += "</tr>";
  }
  // Worksheet名
  const worksheet = "Sheet1";
  const uri = "data:application/vnd.ms-excel;base64,";

  // 下载的表格模板数据
  const template = `<html xmlns:o="urn:schemas-microsoft-com:office:office" 
        xmlns:x="urn:schemas-microsoft-com:office:excel" 
        xmlns="http://www.w3.org/TR/REC-html40">
        <head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>
        <x:Name>${worksheet}</x:Name>
        <x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>
        </x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
        </head><body><table style="table-layout: fixed">${str}</table></body></html>`;
  // 下载模板
  // window.location.href = uri + base64(template);

  const link = document.createElement("a");
  link.href = uri + base64(template);
  // 对下载的文件命名
  link.download = "我是文件名.xls";
  link.click();
};
const base64 = (s) => window.btoa(unescape(encodeURIComponent(s)));

const tableToExcels = () => {
  // 要导出的json数据
  const jsonData = [
    {
      name: "路人甲",
      phone: "123456789",
      email: "000@123456.com",
    },
    {
      name: "炮灰乙",
      phone: "123456789",
      email: "000@123456.com",
    },
    {
      name: "土匪丙",
      phone: "123456789",
      email: "000@123456.com",
    },
    {
      name: "流氓丁",
      phone: "123456789",
      email: "000@123456.com",
    },
  ];
  // 列标题，逗号隔开，每一个逗号就是隔开一个单元格
  let str = `姓名,电话,邮箱\n`;
  // 增加\t为了不让表格显示科学计数法或者其他格式
  for (let i = 0; i < jsonData.length; i++) {
    for (const key in jsonData[i]) {
      str += `${jsonData[i][key] + "\t"},`;
    }
    str += "\n";
  }
  // encodeURIComponent解决中文乱码
  const uri = "data:text/csv;charset=utf-8,\ufeff" + encodeURIComponent(str);
  // 通过创建a标签实现
  const link = document.createElement("a");
  link.href = uri;
  // 对下载的文件命名
  link.download = "json数据表.csv";
  link.click();
};

const exportExcle = () => {
  var data1 = [
    ["id", "name", "value"],
    [1, "sheetjs", 7262],
    [2, "js-xlsx", 6969],
  ];

  var data2 = [
    {
      周一: "语文",
      周二: "数学",
      周三: "历史",
      周四: "政治",
      周五: "英语",
    },
    {
      周一: "数学",
      周二: "数学",
      周三: "政治",
      周四: "英语",
      周五: "英语",
    },
    {
      周一: "政治",
      周二: "英语",
      周三: "历史",
      周四: "政治",
      周五: "数学",
    },
  ];

  //1. 新建一个工作簿
  let workbook = XLSX.utils.book_new();
  //2. 生成一个工作表，
  //2.1 aoa_to_sheet 把数组转换为工作表
  let sheet1 = XLSX.utils.aoa_to_sheet(data1);
  //2.2 把json对象转成工作表
  let sheet2 = XLSX.utils.json_to_sheet(data2);
  //3.在工作簿中添加工作表
  XLSX.utils.book_append_sheet(workbook, sheet1, "sheetName1"); //工作簿名称
  XLSX.utils.book_append_sheet(workbook, sheet2, "sheetName2"); //工作簿名称
  // XLSX.utils.sheet_add_json(sheet1,data2);//把已存在的sheet中数据替换成json数据
  //4.输出工作表,由文件名决定的输出格式
  XLSX.writeFile(workbook, "workBook1.xlsx"); // 保存的文件名
};

const exportcc = () => {
  // STEP 1: Create a new workbook
  const wb = XLSXS.utils.book_new();
  //sheet工作簿标题
  const sheetName = "xlsx导出带样式";
  // STEP 2: Create data rows and styles
  let rowArray = [
    [
      {
        v: "xlsx导出带样式",
        t: "s",
        s: {
          font: {
            name: "Courier",
            sz: 24,
            bold: true,
            color: "b8ddb0",
          },
          fill: { fgColor: { rgb: "9e9e9e" } },
        },
      },
    ],
    [
      {
        v: "Date",
        t: "s",
        s: {
          font: {
            name: "Courier",
          },
        },
      },
      {
        v: "Disburse Amount",
        t: "s",
        s: {
          font: {
            bold: true,
            color: {
              rgb: "FF0000",
            },
          },
        },
      },
      {
        v: "Disburse New Amount",
        t: "s",
        s: {
          fill: {
            fgColor: {
              rgb: "E9E9E9",
            },
          },
        },
      },
      {
        v: "line\nbreak",
        t: "s",
        s: {
          alignment: {
            wrapText: true,
          },
        },
      },
    ],
  ];
  // STEP 3: Create worksheet with rows; Add worksheet to workbook
  const ws = XLSXS.utils.aoa_to_sheet(rowArray);
  XLSXS.utils.book_append_sheet(wb, ws, sheetName);

  let maxColumnNumber = 1; //默认最大列数
  rowArray.map((item) =>
    item.length > maxColumnNumber ? (maxColumnNumber = item.length) : ""
  );
  //合并  #将第一行标题列合并
  let merges = [
    "A1:" + String.fromCharCode(64 + parseInt(maxColumnNumber)) + "1",
  ];
  let wsMerge = [];
  merges.map((item) => {
    wsMerge.push(XLSXS.utils.decode_range(item));
  });

  ws["!merges"] = wsMerge;
  console.log(ws);
  //边框样式
  let borderStyle = {
    top: {
      style: "thin",
      color: {
        rgb: "000000",
      },
    },
    bottom: {
      style: "thin",
      color: {
        rgb: "000000",
      },
    },
    left: {
      style: "thin",
      color: {
        rgb: "000000",
      },
    },
    right: {
      style: "thin",
      color: {
        rgb: "000000",
      },
    },
  };
  //添加外边框
  rowArray.map((item, index) => {
    for (let i = 1; i <= maxColumnNumber; i++) {
      if (!ws["" + String.fromCharCode(64 + parseInt(i)) + (index + 1)]) {
        ws["" + String.fromCharCode(64 + parseInt(i)) + (index + 1)] = {
          v: "",
          t: "s",
          s: { border: borderStyle },
        };
        continue;
      }
      //边框
      ws["" + String.fromCharCode(64 + parseInt(i)) + (index + 1)].s.border =
        borderStyle;

      //字体居中
      ws["" + String.fromCharCode(64 + parseInt(i)) + (index + 1)].s.alignment =
        { vertical: "center", horizontal: "center" };
    }
  });

  //添加列宽
  ws["!cols"] = [
    {
      width: 40,
    },
    {
      width: 40,
    },
  ];
  //添加行高
  ws["!rows"] = [
    { hpt: 40 },
    { hpt: 40 },
    { hpt: 40 },
    { hpt: 40 },
    { hpt: 40 },
    { hpt: 40 },
  ];
  // STEP 4: Write Excel file to browser  #导出
  XLSXS.writeFile(wb, sheetName + ".xlsx");
};
</script>


<style scoped lang="less">
.box {
  width: 80%;
  height: 600px;
  margin: 150px auto;
}
</style>
