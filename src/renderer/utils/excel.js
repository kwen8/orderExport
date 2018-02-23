import XLSX from 'xlsx'

class Excel {
  /**
   * Creates an instance of Excel.
   * @wb 该excel的数据
   * @rABS 是否以数组流的形式读取
   * @memberof Excel
   */
  constructor () {
    this.wb = null
    this.rABS = false
  }

  /**
   * 文件流转BinaryString
   *
   * @param {any} data
   * @returns
   * @memberof Excel
   */
  fixdata (data) {
    let o = ''
    let l = 0
    let w = 10240
    for (; l < data.byteLength / w; ++l) {
      o += String.fromCharCode.apply(
        null,
        new Uint8Array(data.slice(l * w, l * w + w))
      )
    }
    o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)))
    return o
  }

  /**
   * 字符串转字符流
   *
   * @param {any} s
   * @returns
   * @memberof Excel
   */
  s2ab (s) {
    // 字符串转字符流
    let buf = new ArrayBuffer(s.length)
    let view = new Uint8Array(buf)
    for (let i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xff
    return buf
  }

  /**
   * 通过index返回Excel的某个sheet的json数据
   * @param {excel的sheet的索引} index
   */
  getSheetsByIndex (index = 0) {
    return XLSX.utils.sheet_to_json(this.wb.Sheets[this.wb.SheetNames[index]])
  }

  /**
   * 导入excel
   *
   * @param {any} f
   * @param {any} c
   * @memberof Excel
   */
  import (f, c) {
    this.wb = null
    let reader = new FileReader()
    reader.onload = e => {
      const data = e.target.result
      if (this.rABS) {
        this.wb = XLSX.read(btoa(this.fixdata(data)), {
          type: 'base64'
        })
      } else {
        this.wb = XLSX.read(data, {
          type: 'binary'
        })
      }

      if (c && typeof c) {
        c(this.wb)
      }
    }
    if (this.rABS) {
      reader.readAsArrayBuffer(f)
    } else {
      reader.readAsBinaryString(f)
    }
  }

  /**
   * 导入excel
   *
   * @param {any} obj
   * @param {any} c
   * @memberof Excel
   */
  onImport (obj, c) {
    if (!obj.files) {
      return
    }
    this.import(obj.files[0], c)
  }
}

export default new Excel()
