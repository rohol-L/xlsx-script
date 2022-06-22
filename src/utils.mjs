/**
 * 将一个 table 格式的对象分组，转换为 {key:[subTable],……} 的格式
 * @param {Object[]} arr - 需要分组合并的数组。
 * @param {(string|function)} iterateeOrAttr - 按字段名分组或自定义分组
 * @example
 * // returns {'1':[{a:1,b:2},{a:1,b:3}],'2':[{a:2,b:4}]}
 * groupBy([{a:1,b:2},{a:1,b:3},{a:2,b:4}],'a')
 * @returns {Object} 分组结果
 */
function groupBy(arr, iterateeOrAttr) {
  let iteratee = iterateeOrAttr
  if (typeof iterateeOrAttr === 'string') {
    iteratee = (item) => item[iterateeOrAttr]
  }
  if (typeof iteratee !== 'function') {
    throw TypeError('iterateeOrAttr type error.')
  }
  let obj = {}
  arr.forEach((item) => {
    let key = iteratee(item)
    if (obj[key] === undefined) {
      obj[key] = []
    }
    let value = { ...item }
    obj[key].push(value)
  })
  return obj
}

/**
 *
 * @param {Object[]} arr 需要分组的数据
 * @param {(string|function)} iterateeOrAttr - 按字段名分组或自定义分组
 * @param {string=} newColumnName 新的字段名称，默认取分组字段名
 * @param {string=} dataColumnName 新的数据集字段名称，默认为“set”
 * @returns
 */
function groupBy2Arr(arr, iterateeOrAttr, newColumnName, dataColumnName) {
  let objData = groupBy(arr, iterateeOrAttr)
  if (newColumnName === undefined && typeof iterateeOrAttr === 'string') {
    newColumnName = iterateeOrAttr
  }
  return Object.entries(objData).map((en) => ({
    [newColumnName]: en[0],
    [dataColumnName || 'set']: en[1]
  }))
}

/**
 * 取不重复的数据
 * @param {object[]} arr
 * @param {string[]} keys
 * @returns {object[]}
 */
function distinct(arr, keys) {
  let result = []
  for (const row of arr) {
    let find = result.find((r) => keys.every((k) => r[k] === row[k]))
    if (!find) {
      let record = Object.fromEntries(keys.map((k) => [k, row[k]]))
      result.push(record)
    }
  }
  return result
}

/**
 * 取最大值
 * @param {object[]} arr
 * @param {string} keys
 * @returns
 */
function max(arr, key) {
  return arr.map((r) => r[key]).reduce((a, b) => (a > b ? a : b))
}

/**
 * 取最小值
 * @param {object[]} arr
 * @param {string} keys
 * @returns
 */
function min(arr, key) {
  return arr.map((r) => r[key]).reduce((a, b) => (a < b ? a : b))
}

export { groupBy, groupBy2Arr, distinct, max, min }
