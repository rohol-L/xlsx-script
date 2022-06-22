import * as utils from './utils.mjs'

/**
 * 上下文对象
 * @typedef {object} Context
 * @property {string} output 输出
 * @property {object} ws 工作表
 * @property {object} ref_data 引用数据
 * @property {object} exp 表达式
 * @property {Number[]} rc 渲染位置指针
 * @property {Function[]} onRendered 渲染完成后执行
 * @property {boolean} break 终止后续函数计算
 */

const scripts = {
  /**
   * 子模板循环(改变上下文数据集)
   * @param {Context} context 上下文对象
   * @param {Number=} start 循环开始行
   * @param {Number=} end 循环结束行
   */
  for(context, start, end) {
    let { data } = context.ref_data
    const { ws, rc, cell, exp } = context

    if (start === undefined) {
      start = cell.row
    }
    if (end === undefined) {
      end = cell.row
    }

    const rowCount = end - start + 1

    data = utils.groupBy(data, exp.colName)
    const keys = Object.keys(data)
    const length = keys.length

    //本单元格渲染结束后再复制，避免复制到本身
    context.onRendered.push(() => {
      for (let i = 0; i < length; i++) {
        if (i > 0) {
          const copyStart = start + rowCount * i
          this.copyRows(ws, start, end, copyStart)
        }
        const expRowNum = cell.row + rowCount * i
        const key = keys[i]
        this.events.push({
          cell: [expRowNum, cell.col],
          eventName: 'beforeRender',
          callBack: (context) => {
            this.log(2, 'RenderData', data[key])
            context.ref_data.data = data[key]
          }
        })
      }
      //超出范围后恢复数据
      const nextRow = start + rowCount * length
      this.events.unshift({
        cell: [nextRow, 1],
        eventName: 'beforeRender',
        callBack(context) {
          this.log(2, 'RenderData2', data)
          context.ref_data.data = data
        }
      })
    })

    rc[1]--
    context.break = true //中断解析后续公式
  },

  /**
   * 填充行
   * @param {Context} context 上下文对象
   * @param {string} mode 模式：group 分组
   */
  fill(context, mode) {
    const { ws, cell, rc } = context
    let { data } = context.ref_data

    //数据聚合
    if (mode == 'group') {
      let keyInfo = []
      ws.getRow(cell.row).eachCell({ includeEmpty: false }, (cell) => {
        let result = this.parseCell(cell)
        let filterResult = result.exps.filter((exp) => exp.type === null)
        if (filterResult.length == 0) return
        for (const fr of filterResult) {
          keyInfo.push(fr.colName)
        }
      })
      data = utils.distinct(
        data,
        keyInfo.filter((k) => k !== undefined)
      )
    }
    //填充数据
    context.onRendered.push(() => {
      let expCells = []
      ws.getRow(cell.row).eachCell({ includeEmpty: false }, (cell) => {
        let result = this.parseCell(cell)
        let filterResult = result.exps.filter((exp) => exp.type === null)
        if (filterResult.length == 0) return
        expCells.push(result)
      })

      for (let i = 0; i < data.length; i++) {
        if (i > 0) {
          this.copyRows(ws, cell.row, cell.row, cell.row + i)
        }
        let dataRow = data[i]
        let sheetRow = ws.getRow(cell.row + i)

        for (const expc of expCells) {
          let result = []
          for (let i = 0; i < expc.output.length; i++) {
            let exp = expc.exps[i]
            let expResult = ''
            if (exp != null) {
              if (exp.type == null) {
                expResult = dataRow[exp.colName]
              } else {
                expResult = exp.raw
              }
            }
            result.push(expc.output[i], expResult)
          }
          let text = result.join('');
          if (text.length == 0) {
            sheetRow.getCell(expc.cell.col).value = ""
          } else {
            let number = Number(text)
            sheetRow.getCell(expc.cell.col).value = Number.isNaN(number) ? text : number
          }
        }
      }
    })

    rc[1]--
    context.break = true
  },

  /**
   * 过滤最大值(改变上下文数据集)
   * @param {Context} context 上下文对象
   * @param {String=} columnName 字段名
   */
  filterMax(context, columnName) {
    let ref = context.ref_data
    columnName = columnName || context.exp.colName
    let maxValue = utils.max(ref.data, columnName)
    ref.data = ref.data.filter((r) => r[columnName] == maxValue)
  },

  /**
   * 过滤最小值(改变上下文数据集)
   * @param {Context} context 上下文对象
   * @param {String=} columnName 字段名
   */
  filterMin(context, columnName) {
    let ref = context.ref_data
    columnName = columnName || context.exp.colName
    let minValue = utils.min(ref.data, columnName)
    ref.data = ref.data.filter((r) => r[columnName] == minValue)
  },

  /**
   * 选择数据源
   * @param {*} context 
   * @param {*} path 
   */
  select(context, path) {
    let ref = context.ref_data
    let { data, wsData } = ref;

    //history
    if (ref.__selectData === undefined) {
      ref.__selectData = []
    }
    ref.__selectData.push(data)

    let part = path.split("/")
    if (part[0].length == 0) {
      //is absolute path
      data = wsData
    } else if (part[0] == ".") {
      part.unshift();
    }

    for (const p of part) {
      data = data[p]
      if (data === undefined) {
        break;
      }
    }

    ref.data = data;
  },

  /**
   * 撤销选择数据源
   * @param {*} context 
   */
  cancelSelect(context) {
    let ref = context.ref_data
    ref.data = ref.__selectData.pop();
  },

  /**
   * 取数据集的第一项
   * @param {Context} context 上下文对象
   * @param {String=} columnName 字段名
   */
  first(context, columnName) {
    let ref = context.ref_data
    columnName = columnName || context.exp.colName
    context.output = ref.data[0][columnName]
  },

  /**
   * 取数据集的最大值
   * @param {Context} context 上下文对象
   * @param {String=} columnName 字段名
   */
  max(context, columnName) {
    let ref = context.ref_data
    columnName = columnName || context.exp.colName
    context.output = utils.max(ref.data, columnName)
  },

  /**
   * 取数据集的最小值
   * @param {Context} context 上下文对象
   * @param {String=} columnName 字段名
   */
  min(context, columnName) {
    let ref = context.ref_data
    columnName = columnName || context.exp.colName
    context.output = utils.min(ref.data, columnName)
  },

  print(context, text) {
    context.output = text;
  },

  /**
   * 后处理函数 - 合并单元格
   * @param {Context} context 上下文对象
   * @param {*} args
   */
  merge(context, extWidth, extHeight) {
    extWidth = Number(extWidth)
    extHeight = Number(extHeight)
    let { ws, cell } = context
    ws.mergeCells(cell.row, cell.col, cell.row + extHeight, cell.col + extWidth)
  }
}

export default scripts
