/*
 * @Author: ihoey
 * @Date:   2017-11-23 15:07:10
 * @Last Modified by:   ihoey
 * @Last Modified time: 2018-04-17 11:20:15
 */

XLSX = require("xlsx");
json = require("./data");

var title = ['通话Id', '主叫Id', '被叫Id', '是否成功接通', '连接类型', '开始时间', '结束时间', '通话时长（min）', '流量消耗（B）'];
var datas = ['conversationId', 'callUid', 'answerUid', 'callResult', 'connType', 'startTime', 'endTime', 'durationMinutes', 'totalTransfer'];

var _data = json.data.map((e) => {
    tmp = {};
    for (var i = 0; i < title.length; i++) {
        Object.assign(tmp, {
            [title[i]]: e[datas[i]]
        });
    }
    return tmp;
})

var _headers = title;
var headers = _headers
    .map((v, i) => Object.assign({}, { v: v, position: String.fromCharCode(65 + i) + 1 }))
    .reduce((prev, next) => Object.assign({}, prev, {
        [next.position]: { v: next.v }
    }), {});

var data = _data
    .map((v, i) => _headers.map((k, j) => Object.assign({}, { v: v[k], position: String.fromCharCode(65 + j) + (i + 2) })))
    .reduce((prev, next) => prev.concat(next))
    .reduce((prev, next) => Object.assign({}, prev, {
        [next.position]: { v: next.v }
    }), {});

// 合并 headers 和 data
var output = Object.assign({}, headers, data);

// 获取所有单元格的位置
var outputPos = Object.keys(output);

// 计算出范围
var ref = outputPos[0] + ':' + outputPos[outputPos.length - 1];

// 构建 workbook 对象
var wb = {
    SheetNames: ['mySheet'],
    Sheets: {
        'mySheet': Object.assign({}, output, { '!ref': ref })
    }
};

// 导出 Excel
XLSX.writeFile(wb, 'output.xlsx');
