const ExcelJS = require('exceljs')
const fs = require('node:fs/promises')

const INPUT_ANCHOR = '专家姓名'

const INPUT_LAYOUT = [{
  name: '专家姓名',
  x: 0,
  y: 2,
}, {
  name: '身份证号',
  x: 1,
  y: 1,
}, {
  name: '开户银行',
  x: 1,
  y: 3,
}, {
  name: '银行账户',
  x: 3,
  y: 3,
}, {
  name: '手机号',
  x: 5,
  y: 3,
}]

const OUTPUT_LAYOUT = {
  '专家姓名': 2,
  '证件类型': 3,
  '身份证号': 4,
  '国籍': 5,
  '性别': 6,
  '出生日期': 7,
  '手机号': 11,
  '开户银行': 44,
  '银行账户': 45,
}

function plaintext(str) {
  if (!str) return ''
  return str.replaceAll(/[\s\\n\\r\\t]/g, '')
}

// 根据身份证号得到证件类型
function certType(id) {
  if (!id) return ''
  if (id.length != 18) return '?'
  return '居民身份证'
}

// 根据身份证号得到国籍
function nationality(id) {
  if (!id) return ''
  if (id.length != 18) return '?'
  return '中国'
}

// 根据身份证号得到性别
function gender(id) {
  if (!id) return ''
  if (id.length != 18) return '?'
  return (id[16] % 2 == 0) ? '女' : '男'
}

// 根据身份证号得到出生日期
function birthdate(id) {
  if (!id) return ''
  if (id.length != 18) return '?'
  let year = id.substring(6, 10)
  let month = 1 * id.substring(10, 12)
  let date = 1 * id.substring(12, 14)
  return `${year}/${month}/${date}`
}

async function extractFromFile(filename) {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(filename)
  const worksheet = workbook.worksheets[0]

  // 根据数据特征查找到每条数据的位置
  const anchors = Object.keys(worksheet._merges)
    .map(loc => worksheet.getCell(loc))
    .filter(cell => {
      return plaintext(cell.value) === INPUT_ANCHOR
    })

  // 根据指定的数据布局提取出每条数据
  let list = anchors.map(anchor => {
    let x = anchor._column._number
    let y = anchor._row._number

    let data = INPUT_LAYOUT.reduce((data, item) => {
      let c = worksheet.getRow(y + item.y).getCell(x + item.x)
      data[item.name] = plaintext(c.value)
      return data
    }, {})
    return data
  })
    // 抛弃没有姓名的数据
    .filter(data => !!data['专家姓名'])
    .map(data => {
      // 根据身份证号进一步解析相关数据项
      let id = data['身份证号']
      data['证件类型'] = certType(id)
      data['国籍'] = nationality(id)
      data['性别'] = gender(id)
      data['出生日期'] = birthdate(id)
      return data
    })

  return list
}

async function saveResult(result, outputfile) {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('人员信息')
  for (let filename of Object.keys(result)) {
    sheet.addRow([filename])

    let list = result[filename]
    console.log(filename, ':', list.length, '条数据')
    for (let item of list) {
      let rowValues = []
      for (let name of Object.keys(OUTPUT_LAYOUT)) {
        let colNum = OUTPUT_LAYOUT[name]
        rowValues[colNum] = item[name]
      }
      sheet.addRow(rowValues)
    }
  }

  await workbook.xlsx.writeFile(outputfile)
  console.log('已写入文件：', outputfile)
}

async function run() {
  let result = {}

  // 扫描当前文件夹下的所有 .xlsx 文件
  const files = await fs.readdir('./')
  for (const filename of files) {
    if (!filename.endsWith('.xlsx')) continue
    try {
      // 从 .xlsx 文件中提取出数据列表
      let list = await extractFromFile(filename)
      if (list.length > 0) {
        result[filename] = list
      }
    } catch (err) {
      console.error(filename, ':', err.message)
    }
  }

  // 把所有提取到的数据输出到一个 .xlsx 文件
  await saveResult(result, 'output.xlsx')
}

run()
