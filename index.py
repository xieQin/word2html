from docx import Document
from docx.shared import Inches
import webbrowser
import re

title = 'demo'
document = Document(title + '.docx')
GEN_HTML = 'demo.html'

f = open(GEN_HTML,'w')

html = """
<head>
  <meta charset="utf-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
  <meta name="renderer" content="webkit">
  <meta id="meta" name="viewport" content="user-scalable=no, width=device-width, initial-scale=1, maximum-scale=1">
  <title>%s</title>""" %title

html += """
<meta name="description" content="">
  <meta name="keywords" content="">
  <script>
    (window.__setFontSize__ = function() {
      document.documentElement.style.fontSize = Math.min(640, Math.max(document.documentElement.clientWidth, 320)) / 320 * 12 + 'px'
    })()
  </script>
  <style>
    body {
      margin: 0;
    }
    .nk-room {
      counter-reset: sectioncounter;
      font-size: 14px;
      color: #909090;
      line-height: 28px;
      letter-spacing: 1.5px;
      margin: 15px
    }

    .nk-room p {
      text-indent: 28px
    }

    .nk-room p.no-indent {
      text-indent: 0
    }

    .nk-room .no-indent .numb {
      text-align: right;
      padding-right: 45px
    }

    .nk-room p span {
      color: #363636
    }

    .nk-room h4 {
      margin-top: 15px;
      margin-bottom: 15px;
      color: #464646
    }

    .nk-room h5 {
      font-size: 14px;
      margin-top: 15px;
      margin-bottom: 15px
    }

    .nk-room .center {
      text-align: center
    }

    .nk-room strong,.nk-room b {
      color: #363636
    }

    .nk-room table {
      width: 100%;
      border-top: 1px solid #999;
      border-left: 1px solid #999;
      border-spacing: 0;
      color: #909090;
      font-size: 14px
    }

    .nk-room table td {
      width: 65%;
      border-bottom: 1px solid #999;
      border-right: 1px solid #999;
      line-height: 40px
    }

    .nk-room table .tds {
      width: 40%
    }

    .nk-room .left {
      text-indent: 0
    }

    .nk-room .right {
      text-align: right
    }

    .nk-room .mg10 {
      margin-top: .5rem
    }

    .nk-room .f12 {
      font-size: 12px
    }

    .nk-room .in {
      color: #363636
    }

    .nk-room .i100 {
      line-height: 120%;
      letter-spacing: 1px
    }
    .bold {
      font-weight: bold;
    }
    .tc {
      text-align: center;
    }
  </style>
</head>

<body>
<div class='base-box'>
<div class='nk-room'>\n"""

for paragraph in document.paragraphs:
  # print(paragraph.alignment == 1)
  if paragraph.alignment == 1:
    html += "<h4 class='tc mt20 bold'>%s</h4>\n" %paragraph.text
  elif (paragraph.text):
    html += "<p>%s</p>\n" %paragraph.text

html += """
</div>
</div>
</body>
</html>"""

f.write(html)
f.close()

webbrowser.open(GEN_HTML,new = 1)