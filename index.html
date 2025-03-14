<!DOCTYPE html>
<html lang="zh">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>支撑阻力位计算器</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script>
        function calculate() {
            let name = document.getElementById("name").value;
            let high = parseFloat(document.getElementById("high").value);
            let low = parseFloat(document.getElementById("low").value);
            let close = parseFloat(document.getElementById("close").value);
            
            if (isNaN(high) || isNaN(low) || isNaN(close)) {
                alert("请输入有效的数值！");
                return;
            }
            
            let pivot = (high + low + close) / 3;
            let r1 = 2 * pivot - low;
            let r2 = pivot + (high - low);
            let r3 = high + 2 * (pivot - low);
            let s1 = 2 * pivot - high;
            let s2 = pivot - (high - low);
            let s3 = low - 2 * (high - pivot);
            
            document.getElementById("r3").innerText = r3.toFixed(2);
            document.getElementById("r2").innerText = r2.toFixed(2);
            document.getElementById("r1").innerText = r1.toFixed(2);
            document.getElementById("p").innerText = pivot.toFixed(2);
            document.getElementById("s1").innerText = s1.toFixed(2);
            document.getElementById("s2").innerText = s2.toFixed(2);
            document.getElementById("s3").innerText = s3.toFixed(2);
        }

        function exportToExcel() {
            let data = [["名称", "最高价", "最低价", "收盘价", "R3", "R2", "R1", "P", "S1", "S2", "S3"],
                [
                    document.getElementById("name").value,
                    document.getElementById("high").value,
                    document.getElementById("low").value,
                    document.getElementById("close").value,
                    document.getElementById("r3").innerText,
                    document.getElementById("r2").innerText,
                    document.getElementById("r1").innerText,
                    document.getElementById("p").innerText,
                    document.getElementById("s1").innerText,
                    document.getElementById("s2").innerText,
                    document.getElementById("s3").innerText
                ]
            ];
            let ws = XLSX.utils.aoa_to_sheet(data);
            let wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "支撑阻力位");
            XLSX.writeFile(wb, "pivot_points.xlsx");
        }
    </script>
</head>
<body>
    <h2>支撑阻力位计算器</h2>
    <label>名称: <input type="text" id="name"></label><br>
    <label>最高价: <input type="number" id="high" step="0.01"></label><br>
    <label>最低价: <input type="number" id="low" step="0.01"></label><br>
    <label>收盘价: <input type="number" id="close" step="0.01"></label><br>
    <button onclick="calculate()">计算</button>
    <button onclick="exportToExcel()">导出Excel</button>
    <h3>计算结果：</h3>
    <p>R3: <span id="r3"></span></p>
    <p>R2: <span id="r2"></span></p>
    <p>R1: <span id="r1"></span></p>
    <p>P : <span id="p"></span></p>
    <p>S1: <span id="s1"></span></p>
    <p>S2: <span id="s2"></span></p>
    <p>S3: <span id="s3"></span></p>
</body>
</html>
