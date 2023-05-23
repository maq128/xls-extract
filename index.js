const XLSX = require('xlsx')
const fs = require('fs/promises')

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
  '专家姓名': 1,
  '证件类型': 2,
  '身份证号': 3,
  '国籍': 4,
  '性别': 5,
  '出生日期': 6,
  '手机号': 10,
  '开户银行': 43,
  '银行账户': 44,
}

const OUTPUT_FILENAME = 'output.xlsx'

function plaintext(str) {
  if (!str) return ''
  return String(str).replace(/[\s\\n\\r\\t]/g, '')
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
  let workbook = XLSX.readFile(filename)
  let worksheet = workbook.Sheets[workbook.SheetNames[0]]

  let list = []
  for (let loc of Object.keys(worksheet)) {
    if (loc.startsWith('!')) continue
    let v = String(worksheet[loc].v)
    if (plaintext(v) !== INPUT_ANCHOR) continue

    let addr = XLSX.utils.decode_cell(loc)
    let data = INPUT_LAYOUT.reduce((data, item) => {
      let loc = XLSX.utils.encode_cell({ r: addr.r + item.y, c: addr.c + item.x })
      let v = String(worksheet[loc] && worksheet[loc].v || '')
      data[item.name] = plaintext(v)
      return data
    }, {})

    // 抛弃没有姓名的数据
    if (!data['专家姓名']) continue

    // 根据身份证号进一步解析相关数据项
    let id = data['身份证号']
    data['证件类型'] = certType(id)
    data['国籍'] = nationality(id)
    data['性别'] = gender(id)
    data['出生日期'] = birthdate(id)

    list.push(data)
  }
  return list
}

async function saveResult(result, outputfile) {
  let aoa = []
  let row = []
  for (let name of Object.keys(OUTPUT_LAYOUT)) {
    let colNum = OUTPUT_LAYOUT[name]
    row[colNum] = name
  }
  aoa.push(row)

  for (let filename of Object.keys(result)) {
    aoa.push([filename])

    let list = result[filename]
    console.log(`${filename}: ${list.length} 条数据`)
    for (let item of list) {
      let row = []
      for (let name of Object.keys(OUTPUT_LAYOUT)) {
        let colNum = OUTPUT_LAYOUT[name]
        row[colNum] = item[name]
      }
      aoa.push(row)
    }
  }

  let worksheet = XLSX.utils.aoa_to_sheet(aoa)

  let workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, '人员信息')

  XLSX.writeFileXLSX(workbook, outputfile)
  console.log(`已写入文件: ${outputfile}`)
}

async function run() {
  let result = {}

  // 扫描当前文件夹下的所有 .xlsx/.xls 文件
  const files = await fs.readdir('./')
  for (const filename of files) {
    if (!filename.endsWith('.xlsx') && !filename.endsWith('.xls')) continue
    if (filename === OUTPUT_FILENAME) continue
    try {
      // 从 .xlsx/.xls 文件中提取出数据列表
      let list = await extractFromFile(filename)
      if (list.length > 0) {
        result[filename] = list
      }
    } catch (err) {
      console.error(filename, ':', err.message)
    }
  }

  // 把所有提取到的数据输出到一个 .xlsx 文件
  await saveResult(result, OUTPUT_FILENAME)
}

run()
