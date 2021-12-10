import { ChangeEvent, useRef } from "react";
import Excel from "exceljs";

function App() {
  const ref = useRef<HTMLInputElement>(null);

  const handleChange = (e: ChangeEvent) => {
    const files = ref.current?.files || [];
    if (!files.length) return;
    const file = files[0];
    const reader = new FileReader();
    reader.readAsArrayBuffer(file);
    reader.onload = (e) => {
      readBuffer(reader.result as Buffer);
    };
  };

  const readBuffer = async (buffer: Buffer) => {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.getWorksheet(1);
    const obj: { [key: string]: any } = {};
    let names: string[] = [];
    sheet.eachRow((row, i) => {
      let key = "";
      row.eachCell((cell, j) => {
        let value = cell.value?.toString() || "";
        if (i === 1) {
          // 第二列才算数据
          if (j > 1) {
            // 基准行不能为空，空则使用索引
            value = value || `${j - 1}`;
            if (value) {
              obj[value] = {};
              names[j] = value;
            }
          }
        } else {
          if (j === 1) {
            key = value;
          } else {
            obj[names[j]][key] = value;
          }
        }
      });
    });
    for (let key in obj) {
      download(key, obj[key]);
    }
  };

  const download = (filename: string, data: any) => {
    const outBlob = new Blob([JSON.stringify(data)], { type: "text/json" });
    const a = document.createElement("a");
    a.href = window.URL.createObjectURL(outBlob);
    a.download = `${filename}.json`;
    document.body.append(a);
    a.click();
    document.body.removeChild(a);
  };

  return (
    <div className="App">
      <input
        type="file"
        accept=".xlsx,.xls"
        ref={ref}
        onChange={handleChange}
      />
    </div>
  );
}

export default App;
