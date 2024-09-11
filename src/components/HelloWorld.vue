<script setup lang="ts">
import { ref } from 'vue'
import Docxtemplater from 'docxtemplater'
import JSZip from 'pizzip'
import { saveAs } from 'file-saver'
import JSZipUtils from 'jszip-utils'

import { asBlob } from 'html-docx-js-extends'

const handleDocx = () => {
  JSZipUtils.getBinaryContent('./template.docx', function(error, content) {
    // 抛出异常
    if (error) {
      throw error
    }
    // 创建一个JSZip实例，内容为模板的内容
    const zip = new JSZip(content)
    // 创建并加载docxtemplater实例对象
    const doc = new Docxtemplater().loadZip(zip)
    // 设置模板变量的值
    const data = {
      a: '床前明月光，',
      b: '疑是地上霜。',
      c: '举头望明月，',
      d: '低头思故乡。',
      e: '走走走，游游游。',
      f: '甘为铜钱做马牛。',
      g: '做人哪比做妖好。',
      h: '不怕阎王命不休。',
      fors: [
        {
          fa: 'aaa'
        },
        {
          fa: 'bbb'
        },
        {
          fa: 'ccc'
        }
      ]
    }
    doc.setData(data)
    try {
      // 用模板变量的值替换所有模板变量
      doc.render()
    } catch (error) {
      // 抛出异常
      throw error
    }
    // 生成一个代表docxtemplater对象的zip文件（不是一个真实的文件，而是在内存中的表示）
    const out = doc.getZip().generate({
      type: 'blob',
      mimeType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    })
    // 将目标文件对象保存为目标类型的文件，并命名
    saveAs(out, `导出1.docx`)
  })
}

const row = {
  title: '歌曲',
  name: '《半城烟沙》',
  type: 1,
  date: '2024/9/9'
}
const list = [
  {
    a: '有些爱像断线纸鸢',
    b: '结局悲余手中线',
    c: '有些恨像是一个圈',
    d: '冤冤相报不了结',
  },
  {
    a: '只为了完成一个夙愿',
    b: '还将付出几多鲜血',
    c: '忠义之言',
    d: '自欺欺人的谎言',
  },
  {
    a: '有些情入苦难回绵',
    b: '窗间月夕夕成玦',
    c: '有些仇心藏却无言',
    d: '腹化风雪为刀剑',
  },
  {
    a: '只为了完成一个夙愿',
    b: '荒乱中邪正如何辨',
    c: '飞沙狼烟',
    d: '将乱我徒有悲添',
  },
  {
    a: '半城烟沙 兵临池下',
    b: '金戈铁马 替谁争天下',
    c: '一将成 万骨枯',
    d: '多少白发送走黑发',
  },
  {
    a: '半城烟沙 随风而下',
    b: '手中还有 一缕牵挂',
    c: '只盼归田卸甲',
    d: '还能捧回你沏的茶',
  },
]
const getHtml = () => {
  const removeFirst = list.filter((i, index) => {
    if (index !== 0) {
      return i
    }
  })
  let r = ''
  if (list.length) {
    removeFirst.forEach(i => {
      r += `
        <tr>
          <td colspan="1">${i.a}</td>
          <td colspan="1">${i.b}</td>
          <td colspan="1">${i.c}</td>
          <td colspan="1">${i.d}</td>
        </tr>
      `
    })
  }
  return r
}
const handleDocx2 = () => {
      const htmlString =
      `
        <!DOCTYPE html>
        <html lang="en">
        <head>
          <meta charset="UTF-8">
          <title>Document</title>
          <style>
            .main {
              padding: 40px;
              font-family: serif;
              font-weight: 700;
              font-size: 16px;
            }
            .title {
              font-size: 32px;
              text-align: center;
            }
            .title-size {
              font-size: 22px;
            }
            .cert {
              z-index: 3;
            }
            .hcp-table {
              width: 100%;
              margin-top: 30px;
              margin-bottom: 30px;
              border-spacing: 0;
              border-collapse: collapse;
              text-align: center;
              border: 1px solid rgb(12, 12, 12);
            }
            .hcp-table th {
              padding: 8px;
              border: 1px solid rgb(12, 12, 12);
            }
            .hcp-table td {
              padding: 12px 4px;
              border: 1px solid rgb(12, 12, 12);
            }
            .img {
              width: 100px;
              height: 100px;
            }
          </style>
        </head>
        <body>
          <div class="main">
            <div>
              <div class="title">
                ${row.title}
              </div>
              <br/>
              <table class="hcp-table">
                <tr>
                  <td colspan="1" style="width: 100px;padding:4px;" rowspan="2" class="title-size">1</td>
                  <td colspan="2" style="padding:4px;" rowspan="2" class="title-size">2</td>
                  <td colspan="5" style="height:38px; padding:4px;" class="title-size">3</td>
                </tr>
                <tr>
                  <td colspan="1" style="padding: 4px;">4</td>
                  <td colspan="1" style="padding: 4px;">5</td>
                  <td colspan="1" style="padding: 4px;">6</td>
                  <td colspan="1" style="width: 120px; padding:4px;position: relative;z-index: 15;">时间</td>
                  <td colspan="1" style="width: 100px; padding:4px;position: relative;z-index: 15;">
                    头像
                  </td>
                </tr>
                <tr>
                  <td colspan="1" rowspan="${list.length}">${row.name}</td>
                  <td colspan="1">${list[0].a}</td>
                  <td colspan="1">${list[0].b}</td>
                  <td colspan="1" rowspan="${list.length}">${row.type === 1 ? '<div>《如果当时》</div><div>红雨漂泊泛起了回忆怎么潜</div>' : row.type === 2 ? '<div>《如果当时》</div><div>你美目如当年 流转我心间</div>' : ''}
                  </td>
                  <td colspan="1">${list[0].c}</td>
                  <td colspan="1">
                    ${list[0].d}
                  </td>
                  <td ref="imgLeight" colspan="1" rowspan="${list.length}" style="position: relative; z-index: 3">
                    <div class="cert">
                      ${row.date}
                    </div>
                  </td>
                  <td colspan="1" rowspan="${list.length}" style="position: relative; z-index: 9;">
                    <img src="https://cube.elemecdn.com/3/7c/3ea6beec64369c2642b92c6726f1epng.png" alt="" srcset="">
                  </td>
                </tr>
                ${getHtml()}
              </table>
            </div>
            <br/>
          </div>
        </body>
        </html>
      `
      asBlob(htmlString, {
        orientation: 'landscape' // 设置文档方向为横向
      }).then(data => {
        saveAs(data, `导出2.docx`) // save as docx file
      })
}

defineProps<{ msg: string }>()


</script>

<template>
  <h1>{{ msg }}</h1>

  <div class="card">
    <button type="button" @click="handleDocx">导出</button>
  </div>
  <div class="card">
    <button type="button" @click="handleDocx2">表格</button>
  </div>
</template>

<style scoped>
.read-the-docs {
  color: #888;
}
</style>
