import ExcelJS from 'exceljs'
import * as utils from './utils.mjs'
import scripts from './scripts.mjs'
import parseScript from './parser.mjs'

/**
 * 上下文对象
 * @typedef {object} Context
 * @property {string} output 输出
 * @property {object} ws 工作表
 * @property {object} ref_data 引用数据
 * @property {object[]} ref_data.data 上下文数据
 * @property {object[]} ref_data.wsData 工作表数据
 * @property {object} exp 表达式
 * @property {Number[]} rc 渲染位置指针
 * @property {Function[]} onRendered 渲染完成后执行
 * @property {boolean} break 终止后续函数计算
 */

class xlsx_script {
  /**
   * ExcelJS workbook 对象
   */
  workbook

  //单元格事件列表{cell,evName,call}
  events = []

  //函数列表
  scripts = scripts

  //日志开关
  logOutput = 99

  //====渲染相关====//

  /**
   * 渲染模板
   * @param {Object[]} data - 数据源
   */
  render(data) {
    const groupexp = /{#([^}]+)}/

    //eachSheet 不会遍历到新增的sheet
    this.workbook.eachSheet((ws) => {
      const match = ws.name.match(groupexp);
      //普通sheet页直接渲染
      if (match === null) {
        this.renderSheet(ws, data)
        return
      }
      let keyName = ws.name.match(groupexp)[1]
      this.log(2, 'eachSheet', keyName)
      //动态sheet只给支持 {#xxx} 一种语法
      const sheetData = utils.groupBy(data, keyName)
      const count = Object.keys(sheetData).length
      if (count == 0) {
        this.workbook.addWorksheet('无数据').orderNo = ws.orderNo + 1
        this.workbook.removeWorksheet(ws.id)
        return
      }
      //复制n-1个模板
      const sheetList = [ws]
      const entriesData = Object.entries(sheetData)
      ws.name = entriesData[0][0] //原来的模板直接作为第一个待渲染的模板
      for (let i = 1; i < count; i++) {
        sheetList.push(this.copySheet(ws, entriesData[i][0]))
      }

      //渲染生成的模板
      for (let i = 0; i < count; i++) {
        this.renderSheet(sheetList[i], entriesData[i][1])
      }
    })
  }

  /**
   * 渲染worksheet
   * @param {Object} ws - 工作簿
   * @param {Object[]} data - 数据源
   */
  renderSheet(ws, data) {
    //预处理，拆分合并的单元格，避免操作过程中导致合并状态混乱
    this.#preHandleMerge(ws)

    let ref_data = { data, wsData: data }

    //创建上下文对象
    let newContext = (cell, rc) => ({
      output: '',
      ws,
      ref_data,
      exp: null,
      cell,
      rc,
      onRendered: [],
      break: false
    })

    // 检查是否包含指令
    let checkExp = (cell) => {
      return cell != null
        && cell.text
        && cell.text.includes('{')
        && cell.text.includes('}')
    }

    //渲染单元格
    let eachTarget = this.#eachSheetCell(ws, { skipEmpty: false, skipNoCmd: false })
    for (const { cell, rc } of eachTarget) {
      const events = this.#filterEvents('beforeRender', [rc[0] + 1, rc[1] + 1])
      const context = newContext(cell, rc);
      if (events.length > 0) {
        for (const { callBack } of events) {
          callBack && callBack(context)
        }
      }
      if (!checkExp(cell)) continue;
      let parsedCell = this.parseCell(cell)
      this.log(1, 'pos', rc)
      this.renderCell(context, parsedCell, rc)
    }

    //后处理
    eachTarget = this.#eachSheetCell(ws, { skipEmpty: true, skipNoCmd: true })
    for (const { cell } of eachTarget) {
      if (!checkExp(cell)) continue;
      const context = newContext(cell);
      let parsedCell = this.parseCell(cell)
      this.renderCell(context, parsedCell, true)
    }
  }

  /**
   * 渲染单元格
   * @param {Context} context
   * @param {Object} parsedCell
   * @param {string[]} parsedCell.output
   * @param {string} parsedCell.raw
   * @param {Object[]} parsedCell.exps
   * @param {Object} parsedCell.cell
   * @param {boolean} postProcessMode
   */
  renderCell(context, parsedCell, postProcessMode = false) {
    let outputs = []
    for (const exp of parsedCell.exps) {
      if (context.break || exp.type === null || (exp.type == '@') ^ postProcessMode) {
        outputs.push(exp.raw)
        continue
      }
      context.exp = exp
      this.exec(context)
      outputs.push(context.output)
    }
    let combineStr = []
    for (let i = 0; i < parsedCell.output.length; i++) {
      combineStr.push(parsedCell.output[i], outputs[i] || '')
    }
    let value = combineStr.join('');
    if (value === "" || Number.isNaN(Number(value))) {
      parsedCell.cell.value = value
    } else {
      parsedCell.cell.value = Number(value)
    }
    for (const caller of context.onRendered) {
      caller && caller()
    }
  }

  #filterEvents(eventName, cell) {
    return this.events.filter((ev) => ev.eventName == eventName && ev.cell[0] == cell[0] && ev.cell[1] == cell[1])
  }

  /**
   * 执行指令
   * @param {Context} context 
   */
  exec(context) {
    context.output = ''
    let { exp } = context
    for (const func of exp.funcs) {
      if (this.scripts[func.name] === undefined) {
        console.error(func.name + ' 未定义')
        continue
      }
      this.log(2, 'call ' + exp.raw)
      this.scripts[func.name].apply(this, [context, ...func.args])
    }
  }

  //将单元格转换为对象
  parseCell(cell) {
    let raw = cell.text
    let result = {
      output: [],
      raw,
      exps: [],
      cell
    }
    if (raw === null || raw.length == 0) return result

    let { output, exps } = parseScript(raw)
    result.output = output;
    result.exps = exps;

    return result
  }

  //打散合并的单元格，并生成{@.merge(extWidth,extHeigh)}语句
  #preHandleMerge(sheet) {
    for (const model of Object.values(sheet._merges)) {
      let extWidth = model.right - model.left
      let extHeight = model.bottom - model.top
      sheet.unMergeCells(model.tl)
      sheet.getCell(model.tl).value = sheet.getCell(model.tl).text + '{@.merge(' + extWidth + ',' + extHeight + ')}'
    }
  }

  //====Excel操作相关====//

  /**
   * 复制sheet 页
   * @param {Object} templateSheet - 待复制的sheet对象
   * @param {String} name - 复制后的名称
   * @param {Number=} order - 排序位置
   */
  copySheet(templateSheet, name, order) {
    let newSheet = this.workbook.addWorksheet('Sheet')
    if (order === undefined) {
      order = newSheet.orderNo + 1
    }
    newSheet.orderNo = order
    newSheet.model = {
      ...templateSheet.model,
      mergeCells: templateSheet.model.merges
    }
    //修复边框丢失的问题
    for (let i = 0; i < newSheet.model.rows.length; i++) {
      const row = newSheet.model.rows[i]
      const row_t = templateSheet.model.rows[i]
      for (let j = 0; j < row.cells.length; j++) {
        const cell = row.cells[j]
        const cell_t = row_t.cells[j]
        if (Object.keys(cell_t.style).length > 0 && Object.keys(cell.style).length == 0) {
          newSheet.getCell(cell.address).border = { ...cell_t.style.border }
        }
      }
    }
    if (name) {
      newSheet.name = name
    }
    return newSheet
  }

  /**
   * 复制一行
   * @param {Object} sheet - 操作的sheet页
   * @param {Number} templateStart - 复制对象的起始行
   * @param {Number} templateEnd - 复制对象的结束行
   * @param {Number} insetRowNum - 复制后插入的行号
   * @returns void
   */
  copyRows(sheet, templateStart, templateEnd, insetRowNum) {
    if (insetRowNum <= templateEnd) {
      throw RangeError('insetRowNum <= templateEnd')
    }
    let rowCount = templateEnd - templateStart + 1
    let newRows = []
    for (let i = 0; i < rowCount; i++) {
      let t_row = sheet.getRow(templateStart + i)
      let nr = sheet.insertRow(
        insetRowNum + i,
        t_row._cells.map((c) => c.value)
      )
      //复制样式
      t_row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        nr.getCell(colNumber).style = Object.freeze({ ...cell.style })
      })
      nr.height = t_row.heigth
      newRows.push(nr)
    }
    //事件坐标处理
    for (const ev of this.events) {
      if (ev.cell[0] >= insetRowNum) {
        ev.cell[0] += rowCount;
      }
    }
  }

  //迭代所有单元格，返回的 rc 控制当前指针{skipEmpty:true,skipNoCmd:true}
  *#eachSheetCell(sheet, { skipEmpty, skipNoCmd }) {
    let rc = [0, 0] //渲染指针，表示第几行第几列
    for (; rc[0] < sheet._rows.length; rc[0]++) {
      const row = sheet._rows[rc[0]]
      if (row == null) continue
      //下标获取需要+1
      for (rc[1] = 0; rc[1] < row._cells.length; rc[1]++) {
        const cell = row._cells[rc[1]]
        if ((skipEmpty || skipNoCmd) && (cell == null || !cell.text)) {
          continue;
        }
        if (skipNoCmd && !(cell.text.includes('{') && cell.text.includes('}'))) {
          continue;
        }
        yield { cell, rc }
      }
    }
  }

  //====载入数据和导出文件====//

  /**
   * 通过blob/buffer读取文件
   * @param {ArrayBuffer|blob} buffer - 载入blob/buffer
   */
  async loadBuffer(buffer) {
    this.workbook = await new ExcelJS.Workbook().xlsx.load(buffer)
    return this
  }

  /**
   * 通过blob/buffer读取文件
   * @param {ArrayBuffer|blob} buffer - 载入blob/buffer
   */
  static async loadBuffer(buffer) {
    return await new xlsx_script().load(buffer)
  }

  /**
   * 通过url读取文件
   * @param {string} url - 载入url地址
   */
  async loadUrl(url) {
    let rsp = await fetch(url)
    let buffer = await rsp.blob()
    this.workbook = await new ExcelJS.Workbook().xlsx.load(buffer)
    return this
  }

  /**
   * 通过url读取文件
   * @param {string} url - 载入url地址
   */
  static async loadUrl(url) {
    return await new xlsx_script().loadUrl(url)
  }

  /**
   * Node端读取文件
   * @param {*} fileName
   */
  async loadFile(fileName) {
    this.workbook = new ExcelJS.Workbook()
    await this.workbook.xlsx.readFile(fileName)
    return this
  }

  /**
   * Node端读取文件
   * @param {*} fileName
   */
  static async loadFile(fileName) {
    return await new xlsx_script().loadFile(fileName)
  }

  /**
   * 浏览器导出文件
   * @param {String} name - 导出的文件名（不含后缀）
   */
  async export(name) {
    const data = await this.workbook.xlsx.writeBuffer()
    const blob = new Blob([data], {
      type: 'application/octet-stream'
    })
    this.#saveAs(blob, name + '.xlsx')
  }

  /**
   * Node端保存文件
   * @param {string} 文件名
   */
  async save(filename) {
    await this.workbook.xlsx.writeFile(filename)
  }

  #saveAs(blob, fileName) {
    let url = URL.createObjectURL(blob)
    var aLink = document.createElement('a')
    aLink.href = url
    aLink.download = fileName || ''
    var event = new MouseEvent('click')
    aLink.dispatchEvent(event)
  }

  //====其他====//
  /**
   * 日志输出
   * @param {Number} level 日志等级
   * @param {...string} message 日志内容
   */
  log(level, ...message) {
    this.logOutput <= level && console.log(...message)
  }
}

export { xlsx_script }
