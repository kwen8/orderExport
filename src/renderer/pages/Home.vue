<template>
  <div class="drop_area" @dragenter="dragenterHandler" @dragleave="dragleaveHandler" @drop.prevent.stop="dropHandler">
    把文件拖到这里
    <div id="excel"></div>
  </div>
</template>

<script>
import Excel from '../utils/excel'
import Handsontable from 'handsontable'
export default {
  data () {
    return {}
  },
  mounted () {
    const dropArea = document.querySelector('.drop_area')
    if (dropArea) {
      dropArea.addEventListener('dragenter', dragoverHandler, false)
      dropArea.addEventListener('dragover', dragoverHandler, false)
    }
    function dragoverHandler (e) {
      e.stopPropagation()
      e.preventDefault()
      e.dataTransfer.dropEffect = 'copy'
    }
  },
  methods: {
    dragenterHandler (e) {
      e.target.classList.add('active')
    },
    dragleaveHandler (e) {
      e.target.classList.remove('active')
    },
    dropHandler (e) {
      const files = e.dataTransfer

      Excel.onImport(files, () => {
        var rt = Excel.getSheetsByIndex()// <===默认获取Sheet
        var tmp = []
        rt.forEach(function (value, index, array) {
          var t = {}
          for (var k in value) {
            t[k] = value[k]
          }
          tmp.push(t)
        })
        Excel.readDataHead(tmp)
        const hot = new Handsontable(document.getElementById('excel'), {
          data: tmp,
          colHeaders: true,
          rowHeaders: true,
          manualRowMove: true,
          manualColumnMove: true
        })
        console.log(hot.getSourceData())
      })
    }
  }
}
</script>

<style lang="scss">
  .drop_area {
    height: 100%;
    &.active {
      background-color: #cccccc;
    }
  }
</style>
