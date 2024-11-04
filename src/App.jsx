import ExcelJS from "exceljs";
import dayjs from "dayjs";
import { Card, Input, Button, Upload, Space, Form } from "antd";
import { UploadOutlined } from "@ant-design/icons";
import { useEffect, useState } from "react";
// import './App.css'

// index 2 -> productName
// index  12 -> price
const getPrice = (list, product) => {
  const target = list?.find((item) => item?.[2] === product);
  if (target) {
    const price = target[12];
    return price;
  }
  return 0;
};

const download2 = (title, lines, dataList) => {
  // 创建工作簿和工作表
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("产品信息");
  const borderStyle = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };
  // 设置表头
  const style = {
    alignment: { vertical: "middle", horizontal: "center" },
    font: { name: "微软雅黑" },
  };
  worksheet.columns = [
    { header: "序号", key: "index", width: 20, style }, // A
    { header: "商品", key: "product", width: 10, style }, // B
    { header: "单位", key: "quantity", width: 10, style }, // C
    { header: "数量", key: "num", width: 10, style }, // D
    { header: "单价", key: "price", width: 10, style }, // E
    { header: "金额", key: "totalPrice", width: 10, style }, // F
  ];

  const setBorderFATF = (index) => {
    ["A", "B", "C", "D", "E", "F"].forEach((item) => {
      const cell = worksheet.getCell(`${item}${index}`);
      cell.style.border = Object.assign({
        ...cell.style.border,
        ...borderStyle,
      });
    });
  };

  // 添加标题 第一行
  const addTitle = () => {
    worksheet.insertRow(1, {
      index: title,
    });
    worksheet.mergeCells("A1:F1");
    const cell = worksheet.getCell(`A1`);
    cell.style.font = Object.assign({
      ...cell.style.font,
      size: 24,
    });
    // 加边框
    cell.style.border = Object.assign({
      ...cell.style.border,
      ...borderStyle,
      bottom: {},
    });
  };
  addTitle();
  // 添加时间 第二行
  worksheet.insertRow(2, {
    index: dayjs().format("YYYY年MM月DD日"),
  });
  worksheet.mergeCells("A2:F2");
  const cell = worksheet.getCell(`A2`);
  cell.style.alignment = Object.assign({
    ...cell.style.alignment,
    horizontal: "right",
  });
  // 加边框
  cell.style.border = Object.assign({
    ...cell.style.border,
    ...borderStyle,
    top: {},
  });

  const start = 4;
  // 表头增加边框
  setBorderFATF(start - 1);
  // 添加数据行
  lines.map((line, index) => {
    const result = line.match(/([^\d]+)(\d+)(斤)/);
    const [product, num, quantity] = [result[1], result[2], result[3]];
    const price = getPrice(dataList, product);
    worksheet.addRow({
      index: index + 1,
      product,
      quantity,
      num,
      price,
      totalPrice: {
        formula: `=D${index + start}*E${index + start}`,
        result: num * price,
      },
    });
    setBorderFATF(index + start);
  });
  // 添加合并行
  const addMergeRow = () => {
    const end = (lines.length || 0) + start - 1;
    const totalLine = end + 1;
    worksheet.addRow({
      index: "合计",
      totalPrice: {
        formula: `=SUM(F${start}:F${end})`,
      },
    });
    // 合计字体大小
    const cell = worksheet.getCell(`A${totalLine}`);
    cell.style.font = Object.assign({
      ...cell.style.font,
      size: 16,
    });
    setBorderFATF(totalLine);
  };
  addMergeRow();
  // 下载
  workbook.xlsx.writeBuffer().then(function (buffer) {
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    // 创建下载链接
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "product_info.xlsx";
    document.body.appendChild(a);
    a.click();

    // 清理 URL
    URL.revokeObjectURL(url);
  });
};
export default () => {
  const [title, setTitle] = useState("");
  const [dataFile, setDataFile] = useState();
  const [file, setFile] = useState();
  const [list, setList] = useState([]);
  const [dataList, setDataList] = useState([]);
  const getTxt = () => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (event) {
      const str = event.target.result;
      setList(
        str
          ?.split("\n")
          .map((item) => item?.trim())
          .filter(Boolean)
      );
    };
    reader.readAsText(file);
  };
  useEffect(() => {
    getTxt();
  }, [file]);

  useEffect(() => {
    const loadFile = async () => {
      if (!dataFile) return;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(dataFile);
      const worksheet = workbook.getWorksheet(1);
      const sheetValue = worksheet.getSheetValues();
      setDataList(sheetValue);
    };
    loadFile();
  }, [dataFile]);
  return (
    <Card
      style={{ width: "50%" }}
      title={<span>EXCEL 生成工具</span>}
      actions={[
        <Button
          disabled={!file || !dataFile}
          onClick={() => {
            download2(title, list, dataList);
          }}
          type="primary"
        >
          一键生成
        </Button>,
      ]}
    >
      <Form labelCol={{ span: 8 }} wrapperCol={{ span: 16 }}>
        <Form.Item label="表格顶部标题">
          {/* 标题 */}
          <Input
            placeholder="请输入"
            onChange={(e) => setTitle(e.target.value)}
          ></Input>
        </Form.Item>
        <Form.Item label="商品列表文件" required>
          <Upload
            accept=".xlsx,.xlsm"
            beforeUpload={(file) => {
              setDataFile(file);
              return false;
            }}
            onRemove={() => {
              setDataFile(null);
            }}
            maxCount={1}
          >
            <Button icon={<UploadOutlined />} type="dashed">
              请上传文件
            </Button>
          </Upload>
        </Form.Item>
        <Form.Item label="txt格式文件" required>
          {/* 标题 */}
          {/* 文件 .txt */}
          <Upload
            accept=".txt"
            beforeUpload={(file) => {
              setFile(file);
              return false;
            }}
            onRemove={() => {
              setFile(null);
            }}
            maxCount={1}
          >
            <Button icon={<UploadOutlined />} type="dashed">
              请上传文件
            </Button>
          </Upload>
        </Form.Item>
      </Form>
      <Space direction=""></Space>
    </Card>
  );
};
