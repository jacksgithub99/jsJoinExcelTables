
// 导入文件
document.write("<script language = javascript src = 'js/xlsx.core.min.js'></script>")

function testGen() {
    // body...
    xtGenExcel(testData)
}

function testGenStudents() {
    // body...
    xtGenExcel(students, '学生表')
}

function testGenResults() {
    // body...
    xtGenExcel(examResults, '学生成绩表')
}

function xtGenExcel(rows, bookName = 'MYBOOK', sheetName = 'sheet1') {
    // body...
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(rows);

    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

    // 给 sheet 添加【表头】
    // XLSX.utils.sheet_add_aoa(worksheet, [["Name", "Birthday"]], { origin: "A1" });

    // 调整适应列宽
    // const max_width = rows.reduce((w, r) => Math.max(w, r.name.length), 10);
    // 第一列列宽 = max_width, 第二列列宽 = 10。 wch:
    // worksheet["!cols"] = [ { wch: max_width } , {wch: 10}];

    // 输出 excel
    XLSX.writeFile(workbook, bookName + ".xlsx", { compression: true });
}

// 这个方法不需要 xlsx 插件。
function jsonToExcel(jsonData, filename = 'export.xls') {
    if (!Array.isArray(jsonData) || !jsonData?.length) {
        return;
    }
    let str = '';
    // 列标题
    Object.keys(jsonData[0]).forEach(k => {
        str += k + '\t,';
    });
    str += '\n';
    // 增加\t为了不让表格显示科学计数法或者其他格式
    for (let i = 0; i < jsonData.length; i++) {
        // eslint-disable-next-line no-loop-func
        Object.keys(jsonData[i]).forEach(key => {
            str += `${jsonData[i][key] + '\t'},`;
        });
        str += '\n';
    }
    console.log(str)
    // str = '1,2,3,4,\n11,22,33,44,55,66,\na,b,c,d,\nhello'   // \n 换行，也是excel换行。逗号则是代表一列
    // encodeURIComponent解决中文乱码
    const uri = `data:text/${filename.split('.').pop()};charset=utf-8,\ufeff${encodeURIComponent(str)}`;
    console.log(uri)
    // 通过创建a标签实现
    let dom = document.createElement('a');
    dom.download = filename;
    dom.style.display = 'none';
    dom.href = uri;
    document.body.appendChild(dom);
    dom.click();
    setTimeout(() => {
        document.body.removeChild(dom);
    }, 1000);
}

const students = [{"学号":220101,"姓名":"细狗你","班级":1,"性别":"男","年龄":12},{"学号":220102,"姓名":"细鬼","班级":1,"性别":"男","年龄":11},{"学号":220103,"姓名":"渣渣辉","班级":1,"性别":"男","年龄":14},{"学号":220201,"姓名":"吴君如","班级":2,"性别":"女","年龄":12},{"学号":220202,"姓名":"苑琼丹","班级":2,"性别":"女","年龄":13},{"学号":220301,"姓名":"张三","班级":3,"性别":"男","年龄":13},{"学号":220302,"姓名":"王麻子","班级":3,"性别":"女","年龄":12},{"学号":220303,"姓名":"二狗子","班级":3,"性别":"男","年龄":12},{"学号":220304,"姓名":"李四","班级":3,"性别":"男","年龄":13},{"学号":220305,"姓名":"陈华春","班级":3,"性别":"女","年龄":14},{"学号":220306,"姓名":"赵丽华","班级":3,"性别":"女","年龄":12},{"学号":220401,"姓名":"坤哥","班级":4,"性别":"男","年龄":12},{"学号":220402,"姓名":"梁坤","班级":4,"性别":"男","年龄":12},{"学号":220403,"姓名":"梁晨","班级":4,"性别":"男","年龄":12},{"学号":220404,"姓名":"梅静","班级":4,"性别":"女","年龄":12},{"学号":220405,"姓名":"韩汉晨","班级":4,"性别":"男","年龄":13},{"学号":220406,"姓名":"凌楚峰","班级":4,"性别":"男","年龄":12}]
const examResults = [{"学号":220303,"姓名":"二狗子","语文":62,"数学":88,"随机排序":0.534449607369229},{"学号":220301,"姓名":"张三","语文":64,"数学":98,"随机排序":0.166291769352402},{"学号":220101,"姓名":"细狗你","语文":99,"数学":96,"随机排序":0.298748966878335},{"学号":220305,"姓名":"陈华春","语文":83,"数学":94,"随机排序":0.981885816605045},{"学号":220306,"姓名":"赵丽华","语文":87,"数学":97,"随机排序":0.777319924734964},{"学号":220405,"姓名":"韩汉晨","语文":63,"数学":100,"随机排序":0.0758610844747052},{"学号":220403,"姓名":"梁晨","语文":69,"数学":76,"随机排序":0.160749347477528},{"学号":220402,"姓名":"梁坤","语文":52,"数学":83,"随机排序":0.329833118669286},{"学号":220404,"姓名":"梅静","语文":43,"数学":100,"随机排序":0.844087333621201},{"学号":220302,"姓名":"王麻子","语文":67,"数学":66,"随机排序":0.376213770771804},{"学号":220401,"姓名":"坤哥","语文":100,"数学":97,"随机排序":0.736891795092198},{"学号":220304,"姓名":"李四","语文":47,"数学":59,"随机排序":0.801160817943515},{"学号":220406,"姓名":"凌楚峰","语文":55,"数学":98,"随机排序":0.0735291866946264}]

const testData = [
    { name: "George Washington", birthday: "1732-02-22" },
    { name: "John Adams", birthday: "1735-10-19" },
    ]