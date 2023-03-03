

// 导入文件
document.write("<script language = javascript src = 'js/xlsx.full.min.js'></script>")

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

    const worksheet = XLSX.utils.json_to_sheet(rows);        // 无样式  带表头

    // let aoaRows = aoaRowsFrom(rows)
    // console.log(aoaRows)
    // const worksheet = XLSX.utils.aoa_to_sheet(aoaRows);            // 自定义样式，没有表头 （好像TMD仍然没有样式！！！，缺少 xlxs-style?）表头就不能用来拼接

    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

    // // 修改表格中单个单元格的数据（给 sheet 添加【表头】）
    // XLSX.utils.sheet_add_aoa(worksheet, [["Name", "Birthday"]], { origin: "A1" });

    // 调整适应列宽 - 全能
    let fit = true  // true: 自适应列宽； false: 固定列宽
    if(fit) {
        let cws = calculateColumnWidths(rows)
        worksheet["!cols"] = cws
    }
    // let setLaticeStyle = true
    // if(setLaticeStyle){
    //     let lattice = 'C2'
    //     // console.log(lattice)
    //     console.log(latticeStyle())
    //     worksheet['C2'].s.alignment = { vertical: 'center', horizontal: 'center' }//.s = latticeStyle()
    // }
    // 输出 excel
    XLSX.writeFile(workbook, bookName + ".xlsx", { compression: true });
}
function aoaRowsFrom(rows) {
    let aoaRows = []
    for (var i = 0; i < rows.length; i++) {
        let originRow = rows[i]
        let aoaInnerArr = []
        for(k in originRow) {
            let item = {v: originRow[k]}
            item.t = typeof(originRow[k]) == 'number' ? 'n' : 's'
            item.s = {
                font: {
                    bold: true,
                    color: {
                        rgb: "FF0000"
                    }
                }
            }
            aoaInnerArr.push(item)
        }
        aoaRows.push(aoaInnerArr)
    }
    return aoaRows
}
// 计算自适应列宽
function calculateColumnWidths(rows) {
    // !获取表格最终列数及列名！（因为 rows 中可能每一行的列数不同！）
    // 表格的列名，排序优先级：row[0] keys优先级最高，然后是 row[1] 剩余keys,row[2]剩余。
    let coloumKeys = []
    let keyCount = 0
    for (var i = 0; i < rows.length; i++) {
        let row = rows[i]
        for(rk in row)  {
            if(coloumKeys.indexOf(rk) == -1) {
                coloumKeys.push(rk)
            }
        }
    }
    console.log(coloumKeys)
    let cws = []    // column widths
    for (var i = 0; i < coloumKeys.length; i++) {
        cws.push({wch: 10}) // 最小列宽 10 个字符
    }
    for (var i = 0; i < rows.length; i++) {
        let row = rows[i]
        for (var j = 0; j < coloumKeys.length; j++) {
            let rkey = coloumKeys[j]
            if(row[rkey] !== undefined) {
                let str = String(row[rkey])
                let cellW = getCellWidth(str)
                cws[j].wch = Math.max(cws[j].wch, cellW)
            }
        }
    }
    console.log(cws)
    return cws
}
// 字符宽度
function getCellWidth(value) {
  // 版权声明：本文为CSDN博主「云帆Plan」的原创文章，遵循CC 4.0 BY-SA版权协议，转载请附上原文出处链接及本声明。
  // 原文链接：https://blog.csdn.net/a843334549/article/details/114651578
  // 判断是否为null或undefined
  if (value == null) {
    return 10;
  } else if (/.*[\u4e00-\u9fa5]+.*$/.test(value)) {
    // 中文的长度
    const chineseLength = value.match(/[\u4e00-\u9fa5]/g).length;
    // 其他不是中文的长度
    const otherLength = value.length - chineseLength;
    return chineseLength * 2.1 + otherLength * 1.1;
  } else {
    return value.toString().length * 1.1;
    /* 另一种方案
    value = value.toString()
    return value.replace(/[\u0391-\uFFE5]/g, 'aa').length
    */
  }
}


// 使用浏览器自带的href属性导出文件。 1
function jsonToExcel(jsonData, filename = 'from-json.xls') {
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
    stringToExcel(str, filename)
}
// 使用浏览器自带的href属性导出文件。 2
function stringToExcel(astring, filename = 'froms-string.xls') {
    // str = '1,2,3,4,\n11,22,33,44,55,66,\na,b,c,d,\nhello'   // \n 换行，也是excel换行。逗号则是代表一列
    // encodeURIComponent解决中文乱码
    const uri = `data:text/${filename.split('.').pop()};charset=utf-8,\ufeff${encodeURIComponent(astring)}`;
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

const students = [{"学号":220101,"姓名":"细狗你","班级":1,"性别": "男", "年龄": 12, "座右铭": "菩提本无树，明镜亦非台"},{"学号":220102,"姓名":"细鬼","班级":1,"性别":"男","年龄":11},{"学号":220103,"姓名":"渣渣辉","班级":1,"性别":"男","年龄":14, "座右铭": "车到山前必有路"},{"学号":220201,"姓名":"吴君如","班级":2,"性别":"女","年龄":12, "座右铭": "亡羊补牢，为时未晚"},{"学号":220202,"姓名":"苑琼丹","班级":2,"性别":"女","年龄":13, "座右铭": "青春就像卫生纸，看着挺多得，用着用着就不够"},{"学号":220301,"姓名":"张三","班级":3,"性别":"男","年龄":13},{"学号":220302,"姓名":"王麻子","班级":3,"性别":"女","年龄":12},{"学号":220303,"姓名":"二狗子","班级":3,"性别":"男","年龄":12},{"学号":220304,"姓名":"李四","班级":3,"性别":"男","年龄":13},{"学号":220305,"姓名":"陈华春","班级":3,"性别":"女","年龄":14},{"学号":220306,"姓名":"赵丽华","班级":3,"性别":"女","年龄":12},{"学号":220401,"姓名":"坤哥","班级":4,"性别":"男","年龄":12},{"学号":220402,"姓名":"梁坤","班级":4,"性别":"男","年龄":12},{"学号":220403,"姓名":"梁晨","班级":4,"性别":"男","年龄":12},{"学号":220404,"姓名":"梅静","班级":4,"性别":"女","年龄":12},{"学号":220405,"姓名":"韩汉晨","班级":4,"性别":"男","年龄":13},{"学号":220406,"姓名":"凌楚峰","班级":4,"性别":"男","年龄":12, "座右铭": "天高任鸟飞"}]
const examResults = [{"学号":220303,"姓名":"二狗子","语文":62,"数学":88,"随机排序":0.534449607369229},{"学号":220301,"姓名":"张三","语文":64,"数学":98,"随机排序":0.166291769352402},{"学号":220101,"姓名":"细狗你","语文":99,"数学":96,"随机排序":0.298748966878335},{"学号":220305,"姓名":"陈华春","语文":83,"数学":94,"随机排序":0.981885816605045},{"学号":220306,"姓名":"赵丽华","语文":87,"数学":97,"随机排序":0.777319924734964},{"学号":220405,"姓名":"韩汉晨","语文":63,"数学":100,"随机排序":0.0758610844747052},{"学号":220403,"姓名":"梁晨","语文":69,"数学":76,"随机排序":0.160749347477528},{"学号":220402,"姓名":"梁坤","语文":52,"数学":83,"随机排序":0.329833118669286},{"学号":220404,"姓名":"梅静","语文":43,"数学":100,"随机排序":0.844087333621201},{"学号":220302,"姓名":"王麻子","语文":67,"数学":66,"随机排序":0.376213770771804},{"学号":220401,"姓名":"坤哥","语文":100,"数学":97,"随机排序":0.736891795092198},{"学号":220304,"姓名":"李四","语文":47,"数学":59,"随机排序":0.801160817943515},{"学号":220406,"姓名":"凌楚峰","语文":55,"数学":98,"随机排序":0.0735291866946264}]

const testData = [
    { name: "George Washington", birthday: "1732-02-22" },
    { name: "John Adams", birthday: "1735-10-19" },
    ]

function latticeStyle() {
    // body...
    let style = {
        color: {rgb: 'ff0000'},
        fill: { //背景色
                  fgColor: {rgb: '95B3D7'}
               },
        font: {//覆盖字体
                  name: '等线',
                  sz: 10,
                  bold: true
               },

    }
    return style
    
// 版权声明：本文为CSDN博主「en_kai」的原创文章，遵循CC 4.0 BY-SA版权协议，转载请附上原文出处链接及本声明。
// 原文链接：https://blog.csdn.net/en_kai/article/details/128142418
}
