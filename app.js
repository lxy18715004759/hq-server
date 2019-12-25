const express = require('express')
const OSS = require('ali-oss')
const bodyParser = require('body-parser');
const app = express()
const multiparty = require('multiparty')                    // 解析formData文件
const xlsx2json = require("node-xlsx")                      // 解析csv文件
const fs = require('fs')
const path = require('path')
const _ = require('lodash');

const {position, isForeign, areaOfServer, isImportantItem} = require("./config")

function resolve(file) {
    return path.join(__dirname, file)
}

const bucketsMkdir = "hq-datascreen";
const downloadMkdir = "download/"

//配置参数  begin

const businessFields = ["开始日期", "结束日期", "项目指令", "用户名称", "项目负责人", "服务内容", "出差人", "完成情况", "所属科室", "起始位置", "出差位置", "是否国外", "服务地域"]               // 出差CSV文件字段

const businessFieldsOfTest = ["开始日期", "结束日期", "项目指令", "用户名称", "项目负责人", "服务内容", "出差人", "完成情况"]               // 出差CSV文件字段

const hostFields = ["项目指令", "用户名称", "出差位置", "是否国外", "是否重点项目", "机组状态", "服务次数", "服务天数", "服务地域"];
const excelTimeTransformFields = ["开始日期", "结束日期"]

let businessDataOfTest = []         //孵化出的杭汽出差相关的数据
let hostDataOfTest = []           // 孵化出的杭汽机组的相关数据

const defaultStartCity = "浙江省杭州市"
//配置参数  end

//孵化数据
function richData(csvData) {

    if (!Array.isArray(csvData)) {
        console.log("richBusinessData error：csvData is not array");
        return [];
    }
    let concatFields = _.uniq(businessFields.concat(hostFields))        // 整合机组和出差数据

    csvData.map(outItem => {

        let hostName = outItem['用户名称'] ? outItem['用户名称'] : false
        concatFields.map(innerItem => {

            if (!outItem.hasOwnProperty(innerItem)) {

                outItem[innerItem] = "unKnow"

                let endCity = outItem['出差位置'] ? outItem['出差位置'] : false

                if (innerItem === "起始位置") {
                    outItem[innerItem] = defaultStartCity
                }
                if (innerItem === "出差位置") {
                    outItem[innerItem] = hostName ? position[hostName] ? position[hostName] : defaultStartCity : defaultStartCity
                }
                if (innerItem === "所属科室") {
                    outItem[innerItem] = "一科"
                }
                if (innerItem === "是否国外") {
                    outItem[innerItem] = endCity ? isForeign[endCity] ? isForeign[endCity] : false : false
                }
                if (innerItem === "服务地域") {
                    // outItem[innerItem] = endCity ? areaOfServer[endCity] ? areaOfServer[endCity] : areaOfServer[defaultStartCity] : areaOfServer[defaultStartCity]
                    outItem[innerItem] = endCity ? areaOfServer[endCity] ? areaOfServer[endCity] : "国内剩余地区" : "国内剩余地区"
                }
                if (innerItem === "是否重点项目") {
                    outItem[innerItem] = hostName ? isImportantItem[hostName] ? isImportantItem[hostName] : false : false
                }
                if (innerItem === "服务次数") {
                    outItem[innerItem] = 0
                }
                if (innerItem === "服务天数") {
                    outItem[innerItem] = dayDiff(outItem['开始日期'], outItem['结束日期'])
                }

            } else {
                if (!outItem[innerItem]) {
                    outItem[innerItem] = "UnKnow"
                }
            }

        })
        return outItem
    })


    let richBusinessData = _.cloneDeep(csvData)                      // 出差服务数据

    let richHostData = _.cloneDeep(_.uniqBy([...csvData], "项目指令"))

    richHostData.map(outItem => {

        for (let innerKey in outItem) {

            if (innerKey === "服务次数") {
                outItem[innerKey] = _.filter(csvData, {"项目指令": outItem['项目指令']}).length
            }

            if (innerKey === "服务天数") {
                let midHost = _.filter(csvData, {"项目指令": outItem['项目指令']})
                midHost.map(innerItem => {
                    outItem[innerKey] += innerItem['服务天数']
                })
            }
        }
    })

    return {richBusinessData, richHostData};

    function dayDiff(start, end) {
        start = start === "unKnow" ? false : start;
        end = end === "unKnow" ? false : end;

        let startTime = !start ? new Date().getTime() : new Date(start).getTime();
        let endTime = !end ? new Date().getTime() : new Date(end).getTime();
        let maxTime = Math.max(startTime, endTime)
        let minTime = Math.min(startTime, endTime)
        let returnDay = Math.ceil((maxTime - minTime) / (1000 * 3600 * 24))
        return returnDay
    }

}

/**
 * 前端出差显示统计
 * @param data
 * @param type
 */
function statisticsBusinessData(data) {                 // data:csv文件内容    type: host/business

    let statisticsData = {
        endValData: [0, 0, 0],                       // 出差总次数  出差天数  人均出差天数
        locationBusinessChartData: [],       //出差天数
        foreignBusinessChartData: [],        // 出差次数
        locationChartData: [],               // 国内 top
        foreignChartData: [],                // 国外 top
        progressChartData: [],               // 服务内容
        resolveChartData: [],                // 完成情况
        statusChartData: [],                 //科室统计
        areaChartData: [],                   //地域统计
    }
    if (!Array.isArray(data)) {
        console.log("StatisticsBusinessData error: data is not array");
        return false
    }

    //出差次数
    statisticsData.endValData[0] = data.length;

    //出差天数
    let dayFields = ["开始日期", "结束日期"]
    data.map(item => {
        let diff = dayDiff(item[dayFields[0]], item[dayFields[1]])
        item.dayDiff = diff
        statisticsData.endValData[1] += diff
        return item;
    })

    //人均出差天数
    let {businessData, locationTopData, foreignTopData} = businessManData(data)
    let personLength = businessData.length;
    statisticsData.endValData[2] = Math.ceil(statisticsData.endValData[1] / personLength)

    // 出差天数
    let {locationDays, locationDaysPercent, foreignDays, foreignDaysPercent, locationTimes, locationTimesPercent, foreignTimes, foreignTimesPercent} = statisticsLocationAndForeignDaysLength(data)
    statisticsData.locationBusinessChartData = [
        {legend: '国内出差天数', label: "国内", count: locationDays, percent: locationDaysPercent},
        {legend: '国外出差天数', label: "国外", count: foreignDays, percent: foreignDaysPercent}
    ]
    // 出差次数
    statisticsData.foreignBusinessChartData = [
        {legend: '国内出差次数', label: "国内", count: locationTimes, percent: locationTimesPercent},
        {legend: '国外出差次数', label: "国外", count: foreignTimes, percent: foreignTimesPercent}
    ]

    //以个人出差天数为排序标准
    statisticsData.locationChartData = [...locationTopData]

    statisticsData.foreignChartData = [...foreignTopData]


    let {progressData} = progressHandleData(data)
    statisticsData.progressChartData = [...progressData]

    let {resolveData} = resolveHandleData(data)
    statisticsData.resolveChartData = [...resolveData]

    let {statusData} = statusHandleData(businessData)
    statisticsData.statusChartData = [...statusData]

    let {areaData} = areaHandleData(data)
    statisticsData.areaChartData = [...areaData]

    let {geoCoordMap, chinaDatas} = mapHandleData(data);

    statisticsData.mapData = {geoCoordMap, chinaDatas}

    return statisticsData;

    function mapHandleData(data) {
        if (!Array.isArray(data)) {
            console.log("statusHandleData error: data is not array");
            return []
        }

        //起始位置为杭州， 只需要根据出差位置给出坐标位置
        let pointsArr = _.uniqBy(data, "出差位置")
        let positionArr = []                    // 位置数据
        pointsArr.map(item => {

            if (item['出差位置'].match("暂无数据")) {
                item['出差位置'] = "浙江省杭州市"
            }
            if (!positionArr.includes(item['出差位置'])) {
                positionArr.push(item['出差位置'])
            }

        })
        let geoCoordMapConfig = require("./config/location")         //全量位置数据
        // 匹配到的位置信息
        let chinaDatas = [
            [
                {
                    name: "总部",
                    value: 0
                }
            ]
        ]
        let geoCoordMap = {
            总部: ["120.134933", "30.29459"]
        }
        positionArr.map(item => {
            if (!geoCoordMap.hasOwnProperty(item)) {
                geoCoordMap[item] = geoCoordMapConfig[item] ? geoCoordMapConfig[item] : geoCoordMapConfig["浙江省杭州市"]
                let arr = [{name: item, value: 0}]
                chinaDatas.push(arr)
            }
        })
        return {chinaDatas, geoCoordMap}
    }

    function statusHandleData(personObj) {

        if (!Array.isArray(data)) {
            console.log("statusHandleData error: data is not array");
            return []
        }

        let personArr = []
        Object.keys(personObj).forEach(key => {
            personArr.push(personObj[key])
        })

        personArr = randomClassRoom(personArr)

        let onlineOfClassRoomOne = _.filter(personArr, {"所属科室": "一科", "status": "online"}).length
        let hideOfClassRoomOne = _.filter(personArr, {"所属科室": "一科", "status": "hide"}).length

        let onlineOfClassRoomTwo = _.filter(personArr, {"所属科室": "二科", "status": "online"}).length
        let hideOfClassRoomTwo = _.filter(personArr, {"所属科室": "二科", "status": "hide"}).length

        let onlineOfClassRoomThree = _.filter(personArr, {"所属科室": "三科", "status": "online"}).length
        let hideOfClassRoomThree = _.filter(personArr, {"所属科室": "三科", "status": "hide"}).length

        let onlineOfClassRoomFour = _.filter(personArr, {"所属科室": "四科", "status": "online"}).length
        let hideOfClassRoomFour = _.filter(personArr, {"所属科室": "四科", "status": "hide"}).length

        let statusData = [
            {
                class: "出差",
                "一科": hideOfClassRoomOne,
                "二科": hideOfClassRoomTwo,
                "三科": hideOfClassRoomThree,
                "四科": hideOfClassRoomFour
            },
            {
                class: "在杭",
                "一科": onlineOfClassRoomOne,
                "二科": onlineOfClassRoomTwo,
                "三科": onlineOfClassRoomThree,
                "四科": onlineOfClassRoomFour
            }
        ]

        return {statusData}

    }

    function randomClassRoom(data) {

        let a = Math.ceil(data.length / 4)
        data.map((item, key) => {
            if (key < a) {
                item["所属科室"] = "一科"
            } else if (key < 2 * a) {
                item["所属科室"] = "二科"
            } else if (key < 3 * a) {
                item["所属科室"] = "三科"
            } else {
                item["所属科室"] = "四科"
            }
            item['status'] = randomStatus()
            return item
        })
        return data

        function randomStatus() {
            return Math.random() < 0.5 ? "online" : "hide"
        }


    }

    function areaHandleData(data) {

        if (!Array.isArray(data)) {
            console.log("areaHandleData error: data is not array");
            return []
        }

        let areaData = []

        let a = _.filter(data, {"服务地域": "新疆&西北地区"}).length
        let b = _.filter(data, {"服务地域": "东北&西南地区"}).length
        let c = _.filter(data, {"服务地域": "国外"}).length
        let d = _.filter(data, {"服务地域": "国内剩余地区"}).length

        areaData.push({area: "新疆&西北地区", value: a}, {area: "东北&西南地区", value: b}, {area: "国外", value: c}, {
            area: "国内剩余地区",
            value: d
        })
        return {areaData}

    }

    function resolveHandleData(data) {

        if (!Array.isArray(data)) {
            console.log("resolveHandleData error: data is not array");
            return []
        }

        let finish = _.filter(data, {"完成情况": "完成"}).length
        let unfinish = _.filter(data, {"完成情况": "未完成"}).length
        let others = _.filter(data, {"完成情况": "UnKnow"}).length
        const cls = {
            "已处理": finish,
            "处理中": others,
            "未处理": unfinish
        }
        let total = data.length

        let resolveData = []
        for (let a in cls) {
            resolveData.push({
                "item": a,
                "count": cls[a],
                "percent": Number((cls[a] / total + "").substr(0, 4)),
                "showItem": false
            })
        }

        return {resolveData}
    }

    function progressHandleData(data) {

        if (!Array.isArray(data)) {
            console.log("progressHandleData error: data is not array");
            return []
        }

        let cls = {
            "派员交底": 0,
            "安装指导": 0,
            "调试指导": 0,
            "问题处理": 0,
            "商务出差": 0,
        }
        let clsRules = {
            "派员交底": ["技术交流", "技术交底", "培训", "安装交底", "技术指导"],
            "安装指导": ["安装指导", "安装调试", "安装", "安装服务", "仪表指导"],
            "调试指导": ["试车指导", "调试", "调试指导", "试验", "试航", "远程监测", "试车"],
            "问题处理": ["大修", "问题处理", "送505", "性能测试", "送货", "返修", "上海大隆配短"],
            "商务出差": ["促销", "用户走访", "unKnow", "用户回访", "应收账款", "巡访", "用户巡访", "回访", "面签", "A检"],
        }
        data.map(item => {
            let mid = item['服务内容'].trim()

            for (let innerKey in clsRules) {

                if (clsRules[innerKey].includes(mid)) {
                    cls[innerKey]++
                }
            }

        })
        let progressData = []
        for (let a in cls) {
            progressData.push({
                serviceItem: a,
                count: cls[a]
            })
        }

        return {progressData}
    }

    //统计 国内外天数
    function statisticsLocationAndForeignDaysLength(data) {

        let locationDays = 0;
        let locationDaysPercent = 0;
        let foreignDays = 0;
        let foreignDaysPercent = 0;

        let locationTimes = 0;
        let locationTimesPercent = 0;
        let foreignTimes = 0
        let foreignTimesPercent = 0

        data.map(outItem => {
            if (outItem['是否国外']) {
                foreignTimes += 1
                foreignDays += outItem.dayDiff
            } else {
                locationTimes += 1
                locationDays += outItem.dayDiff
            }
        })

        let totalDays = locationDays + foreignDays
        let totalTimes = data.length;

        locationDaysPercent = Number(((locationDays / totalDays) + "").slice(0, 4))
        foreignDaysPercent = Number(((1 - locationDaysPercent + "").slice(0, 4)))

        locationTimesPercent = Number(((locationTimes / totalTimes) + "").slice(0, 4))
        foreignTimesPercent = Number(((1 - locationTimesPercent) + "").slice(0, 4))

        return {
            locationDays,
            foreignDays,
            locationDaysPercent,
            foreignDaysPercent,
            locationTimes,
            locationTimesPercent,
            foreignTimes,
            foreignTimesPercent
        }
    }

    /**
     * 计算每个人的出差次数、天数
     * @param data
     */
    function businessManData(data) {
        let cls = {}
        data.map(item => {

            let midArr = item['出差人'] ? item['出差人'].split("、") : []
            let day = dayDiff(item['开始日期'], item['结束日期'])

            if (item['是否国外']) {

                midArr.map(innerItem => {
                    if (cls[innerItem]) {
                        cls[innerItem]['foreignTime'] += 1
                        cls[innerItem]['foreignDayDiff'] += day
                    } else {
                        cls[innerItem] = {locationTime: 0, locationDayDiff: 0, foreignTime: 0, foreignDayDiff: 0}
                        cls[innerItem]['foreignTime'] = 1
                        cls[innerItem]['foreignDayDiff'] = day
                        cls[innerItem]["所属科室"] = item['所属科室']
                    }
                })
            } else {

                midArr.map(innerItem => {
                    if (cls[innerItem]) {
                        cls[innerItem]['locationTime'] += 1
                        cls[innerItem]['locationDayDiff'] += day
                    } else {
                        cls[innerItem] = {locationTime: 0, locationDayDiff: 0, foreignTime: 0, foreignDayDiff: 0}
                        cls[innerItem]['locationTime'] = 1
                        cls[innerItem]['locationDayDiff'] = day
                        cls[innerItem]["所属科室"] = item['所属科室']
                    }
                })
            }

        })
        let businessData = []
        for (let a in cls) {
            businessData.push({
                "name": a || "unKnow",
                "locationTime": cls[a]['locationTime'],
                "locationDayDiff": cls[a]["locationDayDiff"],
                "foreignTime": cls[a]['foreignTime'],
                "foreignDayDiff": cls[a]["foreignDayDiff"],
                "所属科室": cls[a]['所属科室']
            })
        }

        function joinString(arr, timeField, dayDiffField) {

            if (!Array.isArray(arr)) {
                console.log("joinString error : arr is not array");
                return []
            }

            let returnArr = []
            arr.map(item => {

                let midArr = ["", "", '']
                midArr[0] = item['所属科室'] + "-" + item['name']
                midArr[1] = "出差" + item[timeField] + "次"
                midArr[2] = "出差" + item[dayDiffField] + "天"
                returnArr.push(midArr)
            })
            return returnArr
        }

        let locationTopData = joinString(_.orderBy(businessData, ["locationDayDiff"], ["desc"]).slice(0, 6), "locationTime", "locationDayDiff")
        let foreignTopData = joinString(_.orderBy(businessData, ["foreignDayDiff"], ['desc']).slice(0, 6), "foreignTime", "foreignDayDiff")

        return {businessData, locationTopData, foreignTopData};
    }
}

/**
 * 计算时差
 * @param start
 * @param end
 * @returns {number}
 */
function dayDiff(start, end) {
    start = start === "UnKnow" ? false : start;
    end = end === "UnKnow" ? false : end;

    let startTime = !start ? new Date().getTime() : new Date(start).getTime();
    let endTime = !end ? new Date().getTime() : new Date(end).getTime();
    let maxTime = Math.max(startTime, endTime)
    let minTime = Math.min(startTime, endTime)

    let returnDay = Math.ceil((maxTime - minTime) / (1000 * 3600 * 24))
    return returnDay
}

function statisticsHostData(data) {

    let statisticsData = {
        endValData: [0, 0, 0],
        locationBusinessChartData: [],        //
        foreignBusinessChartData: [],
        locationChartData: [],
        foreignChartData: [],
        resolveChartData: [],
        areaChartData: [],
    }
    if (!Array.isArray(data)) {
        console.log("statisticsHostData error: data is not array");
        return false
    }

    let locationImportantHostArr = _.filter(data, {"是否重点项目": true})
    let foreignImportantHostArr = _.filter(data, {"是否重点项目": false})

    statisticsData.endValData[0] = data.length;
    statisticsData.endValData[1] = locationImportantHostArr.length
    statisticsData.endValData[2] = foreignImportantHostArr.length

    let locationHostArr = _.filter(data, {"是否国外": false})
    let foreignHostArr = _.filter(data, {"是否国外": true})

    statisticsData.locationBusinessChartData = [
        {
            legend: '国内机组数量',
            label: "国内",
            count: locationHostArr.length,
            percent: Number((locationHostArr.length / data.length + "").slice(0, 4))
        },
        {
            legend: '国外机组数量',
            label: "国外",
            count: foreignHostArr.length,
            percent: Number((foreignHostArr.length / data.length + "").slice(0, 4))
        }
    ]

    let locationImportantArr = _.filter(data, {"是否国外": false, "是否重点项目": true})
    let foreignImportantArr = _.filter(data, {"是否国外": true, "是否重点项目": true})

    statisticsData.foreignBusinessChartData = [
        {
            legend: '国内重点项目',
            label: "国内",
            count: locationImportantArr.length,
            percent: Number((locationImportantArr.length / locationImportantHostArr.length + "").slice(0, 4))
        },
        {
            legend: '国外重点项目',
            label: "国外",
            count: foreignImportantArr.length,
            percent: Number((foreignImportantArr.length / locationImportantHostArr.length + "").slice(0, 4))
        }
    ]

    let locationTopData = joinString(_.orderBy(locationHostArr, ["服务天数"], ["desc"]).slice(0, 6))
    let foreignTopData = joinString(_.orderBy(foreignHostArr, ["服务天数"], ["desc"]).slice(0, 6))

    statisticsData.locationChartData = [...locationTopData]
    statisticsData.foreignChartData = [...foreignTopData]

    statisticsData.resolveChartData = [
        {item: '安装阶段', count: 392, percent: 0.16, showItem: true},
        {item: '调试阶段', count: 204, percent: 0.12, showItem: true},
        {item: '问题处理', count: 143, percent: 0.7, showItem: true},
        {item: '已完成', count: 2000, percent: 0.72, showItem: true}
    ]

    statisticsData.areaChartData = [
        {area: "新疆&西北地区", value: _.filter(data, {"服务地域": "新疆&西北地区"}).length},
        {area: "东北&西南地区", value: _.filter(data, {"服务地域": "东北&西南地区"}).length},
        {area: "国内剩余地区", value: _.filter(data, {"服务地域": "国内剩余地区"}).length},
        {area: "国外", value: _.filter(data, {"服务地域": "国外"}).length}
    ]
    let {geoCoordMap, chinaDatas} = mapHandleData(data)
    statisticsData.mapData = {
        geoCoordMap,
        chinaDatas
    }

    return statisticsData;

    function mapHandleData(data) {
        if (!Array.isArray(data)) {
            console.log("statusHandleData error: data is not array");
            return []
        }

        //起始位置为杭州， 只需要根据出差位置给出坐标位置
        let pointsArr = _.uniqBy(data, "出差位置")
        let positionArr = []                    // 位置数据
        pointsArr.map(item => {
            if (item['出差位置'].match("暂无数据")) {
                item['出差位置'] = "浙江省杭州市"
            }
            if (!positionArr.includes(item['出差位置'])) {
                positionArr.push(item['出差位置'])
            }

        })
        let geoCoordMapConfig = require("./config/location")         //全量位置数据
        // 匹配到的位置信息
        let chinaDatas = [
            [
                {
                    name: "总部",
                    value: 0
                }
            ]
        ]
        let geoCoordMap = {
            总部: ["120.134933", "30.29459"]
        }
        positionArr.map(item => {
            if (!geoCoordMap.hasOwnProperty(item)) {
                geoCoordMap[item] = geoCoordMapConfig[item] ? geoCoordMapConfig[item] : geoCoordMapConfig["浙江省杭州市"]
                let arr = [{name: item, value: 0}]
                chinaDatas.push(arr)
            }
        })
        return {chinaDatas, geoCoordMap}
    }

    function joinString(arr) {

        if (!Array.isArray(arr)) {
            console.log("joinString error : arr is not array");
        }
        let rArr = []
        arr.map(item => {
            let midArr = ["", "", '']
            midArr[0] = item['项目指令']
            midArr[1] = "服务" + item['服务次数'] + "次"
            midArr[2] = "服务" + item['服务天数'] + "天"
            rArr.push(midArr)
        })

        return rArr;
    }
}

//读取本地文件
function readDir(path, type) {

    return new Promise((resolve, reject) => {

        if (!path) {
            reject()
        }
        fs.readdir(path, (err, files) => {
            if (!err) {
                if (files.length === 0) {
                    reject(`请上传${type}文件`)
                }
                resolve(files);
            } else {
                reject(`服务器错误：读取${type}文件错误！`)
            }
        })

    })

}

//
/**
 * 筛选出本地最新的文件
 * @param files
 * @param type
 * @param timeIndex             兼容本地读取文件的命名格式以及OSS上的命名格式
 * @returns {boolean|*}
 */
function filterDir(files, type, timeIndex = 1) {                              // type : host  business video    files: 文件格式：business-timestamp.xlsx   host-timestamp.xlsx

    if (!Array.isArray(files)) {
        console.log("FilterDir error: files is not array")
        return false;
    }
    if (!type) {
        console.log("FilterDir error: type is not defined")
        return false;
    }

    let maxTimeFile = "";         // 时间戳最大的文件名
    let maxTimestamp = 0


    if (files[0]) {
        let fileNameArr = files[0].split('-')           // 文件名以“-”分割成的数组
        if (fileNameArr[0] === type) {
            maxTimeFile = files[0];         // 时间戳最大的文件名
            maxTimestamp = parseInt(files[0].split("-")[timeIndex].split(".")[0])
        }
    }

    files.map(item => {
        let fileNameArr = item.split('-')           // 文件名以“-”分割成的数组
        if (fileNameArr[0] === type) {
            // 具有相同文件名前缀的文件  host
            let timeArr = fileNameArr[timeIndex].split(".")         // 文件名以“.”分割成的数据

            if (parseInt(timeArr[0]) > maxTimestamp) {

                maxTimestamp = timeArr[0]
                maxTimeFile = item;
            }
        }
    })
    if (maxTimeFile.length === 0) {
        let msg = timeIndex === 1 ? `filterDir error:本地暂无${type}文件,请先下载到本地。` : `filterDir error:服务器暂无${type}文件,请先上传本地文件至服务器。`
        console.log(msg);
    }
    return maxTimeFile;

}


app.use(bodyParser.json());//数据JSON类型
app.use(bodyParser.urlencoded({extended: false}));//解析post请求数据

let client = new OSS({
    // region: '<oss region>',
    //云账号AccessKey有所有API访问权限，建议遵循阿里云安全最佳实践，部署在服务端使用RAM子账号或STS，部署在客户端使用STS。
    accessKeyId: 'LTAI4FjtDa1VvvbXLWHqe3fP',
    accessKeySecret: 'pO87v50FWjUALuh0jA3nS7ZayA7X23',
    bucket: bucketsMkdir
})

// 查看所有 Bucket
async function listBuckets(client) {
    try {
        const result = await client.listBuckets();
        let bucketsList = [...result.buckets]
        return bucketsList;
    } catch (err) {
        console.log(err);
    }
}

//查看指定目录下的文件
async function listFilesByBucketName(client, bucketName, prefix = "") {

    client.useBucket(bucketName)
    try {
        const result = await client.list({
            'max-keys': 5,
            prefix
        })
        return result;
    } catch (e) {
        console.log(e);
    }
}

// 删除 Bucket
async function deleteBucket(client, bucketName) {
    try {
        const result = await client.deleteBucket(bucketName);
        console.log(result);
    } catch (err) {
        console.log(err);
    }
}

// 分片上传文件
async function multipartUpload(client, objectName, localFile) {

    try {
        client.useBucket(bucketsMkdir)
        let result = await client.multipartUpload(objectName, localFile, {
            meta: {
                year: 2017,
                people: 'test'
            }
        });
        let head = await client.head(objectName);
        return result;
    } catch (e) {
        // 捕获超时异常
        if (e.code === 'ConnectionTimeoutError') {
            console.log("Woops,超时啦!");
            // do ConnectionTimeoutError operation
        }
        return false;
    }
}

// 下载到本地文件
//读取到内存中
async function getToLocal(client, objectName, localFile) {

    try {
        let result = await client.get(objectName);
        return result;
    } catch (e) {
        console.log(e);
    }
}

//下载到本地
async function downToLocal(client, objectName, localFile) {

    try {
        let result = await client.getStream(objectName);
        let writeStream = fs.createWriteStream(localFile);
        result.stream.pipe(writeStream);
    } catch (e) {
        console.log(e);
    }
}

// excel读取2018/01/01这种时间格式是会将它装换成数字类似于46254.1545151415 numb是传过来的整数数字，format是之间间隔的符号
function excelTimeTransform(numb) {
    const time = new Date(numb * 24 * 3600000 + 1)
    time.setYear(time.getFullYear() - 70)
    const year = time.getFullYear() + ''
    const month = time.getMonth() + 1 + ''
    const date = time.getDate() - 1 + ''

    return year + "-" + (month < 10 ? '0' + month : month) + "-" + (date < 10 ? '0' + date : date)
}

//读取本地CSV文件
function getLocalFile(file, csvFields) {
    return new Promise((resolve) => {
        let list = xlsx2json.parse(file);
        let data = [...(list[0].data)];
        let arr = [];
        let title = data[0]    // 表头
        //去除括号
        let tableTitle = []
        title.map(item => {
            let midTitle = ""
            let mat = item.match(/\(/)
            if (mat) {
                let matInner = item.match(/\（/)
                if (matInner) {
                    let index = item.indexOf("（")
                    midTitle = item.slice(0, index)
                } else {
                    let index = item.indexOf("(")
                    midTitle = item.slice(0, index)
                }
            } else {
                midTitle = item;
            }

            tableTitle.push(midTitle)
        })
        let dataArr = data.slice(1)
        dataArr.map(outItem => {
            let params = {}
            outItem.map((innerItem, innerKey) => {

                if (excelTimeTransformFields.includes(tableTitle[innerKey])) {
                    params[tableTitle[innerKey]] = excelTimeTransform(innerItem)
                } else {
                    params[tableTitle[innerKey]] = innerItem
                }
            })

            //填充完整数据，null填充
            csvFields.map(item => {
                if (!params.hasOwnProperty(item)) {
                    params[item] = null;
                }
            })
            arr.push(params)
        })

        resolve(arr)
    })
}

const addZero = (num) => {
    if (isNaN(num)) return
    return num <= 9 ? '0' + num : num + ''
}

const dateFormatter = (stamptime) => {
    const datetime = new Date(stamptime)
    return datetime.getFullYear() + '-' + addZero(datetime.getMonth() + 1) + '-' + addZero(datetime.getDate()) + ' ' + addZero(datetime.getHours()) + ':' + addZero(datetime.getMinutes()) + ':' + addZero(datetime.getSeconds())
}

function getLocalFileOfTest(file, csvFields) {
    return new Promise((resolve) => {
        let list = xlsx2json.parse(file);
        let data = [...(list[0].data)];
        let arr = [];
        let title = data[0]    // 表头
        //去除括号
        let tableTitle = []
        title.map(item => {
            let midTitle = ""
            let mat = item.match(/\(/)
            if (mat) {
                let matInner = item.match(/\（/)
                if (matInner) {
                    let index = item.indexOf("（")
                    midTitle = item.slice(0, index)
                } else {
                    let index = item.indexOf("(")
                    midTitle = item.slice(0, index)
                }
            } else {
                midTitle = item;
            }

            tableTitle.push(midTitle)
        })
        let dataArr = data.slice(1)
        dataArr.map(outItem => {
            let params = {}
            outItem.map((innerItem, innerKey) => {

                if (excelTimeTransformFields.includes(tableTitle[innerKey])) {

                    params[tableTitle[innerKey]] = dateFormatter(innerItem).split(" ")[0]
                } else {
                    params[tableTitle[innerKey]] = innerItem
                }
            })

            //填充完整数据，null填充
            csvFields.map(item => {
                if (!params.hasOwnProperty(item)) {
                    params[item] = null;
                }
            })
            arr.push(params)
        })

        resolve(arr)
    })
}


app.use("*", function (req, res, next) {
    res.header("Access-Control-Allow-Origin", "*")
    res.header("Access-Control-Allow-Headers", "X-Requested-With, mytoken")
    res.header("Access-Control-Allow-Headers", "X-Requested-With, Authorization")
    res.setHeader('Content-Type', 'application/json;charset=utf-8')
    res.header("Access-Control-Allow-Headers", "Content-Type,Content-Length, Authorization, Accept,X-Requested-With");
    res.header("Access-Control-Allow-Methods", "PUT,POST,GET,DELETE,OPTIONS");
    res.header("X-Powered-By", ' 3.2.1')
    if (req.method == "OPTIONS") res.send(200);/*让options请求快速返回*/
    else next()
})

app.get("/getBuckets", (req, res) => {

    let returnData = {
        code: 200,
        data: [],
        msg: ""
    }
    listBuckets(client).then(_ => {
        console.log('%c listBuckets' + _, "color:green");
        ;
        returnData.data = _
        res.send(returnData)
    }).catch(err => {
        console.log('%c listBuckets' + err, "color:red");
        returnData.code = 0
        returnData.msg = "error"
        res.send(returnData)
    })

})

app.get("/getFiles", (req, res) => {
    let bucketName = bucketsMkdir
    let returnData = {
        code: 200,
        data: [],
        msg: ""
    }
    listFilesByBucketName(client, bucketName).then(_ => {
        returnData.data = _.objects
        res.send(returnData)
    }).catch(err => {
        returnData.code = 0
        returnData.msg = "error"
        res.send(returnData)
    })

})

const separator = "&"         // 文件之间的分割符
//上传至OSS服务器（分片上传）   .xlsx  .mp4
app.post("/uploadFile", (request, response) => {
    let form = new multiparty.Form()
    form.parse(request, (err, fields, files) => {
        let fileName = fields.fileName[0]
        let returnData = {
            code: 200,
            data: '',
            msg: "上传成功！"
        }
        let fullName = files.file[0].originalFilename
        let prefixName = fullName.split(".")[0]
        let fileType = fullName.split(".")[1]
        let objectName = fileName + separator + prefixName + separator + new Date().getTime() + "." + fileType;                 // 上传文件命名
        let localFile = files.file[0].path;

        multipartUpload(client, objectName, localFile).then(res => {
            returnData.code = 200;
            returnData.msg = "success"
            response.send(returnData);
        }).catch(err => {
            returnData.code = 0;
            returnData.msg = "error"
            response.send(returnData);
        })

    })
})

/**
 * 转码buffer数据
 * @param arr
 * @returns {string}
 */
function byteToString(arr) {
    if (typeof arr === 'string') {
        return arr;
    }
    let str = '',
        _arr = arr;
    for (let i = 0; i < _arr.length; i++) {
        let one = _arr[i].toString(2),
            v = one.match(/^1+?(?=0)/);
        if (v && one.length == 8) {
            let bytesLength = v[0].length;
            let store = _arr[i].toString(2).slice(7 - bytesLength);
            for (let st = 1; st < bytesLength; st++) {
                store += _arr[st + i].toString(2).slice(2);
            }
            str += String.fromCharCode(parseInt(store, 2));
            i += bytesLength - 1;
        } else {
            str += String.fromCharCode(_arr[i]);
        }
    }
    return str;
}

//从OSS服务下载最新文件
app.post("/upDownFile", (request, response) => {

    let returnData = {
        code: 0,
        data: [],
        msg: ""
    }

    let {fileType} = request.body

    listFilesByBucketName(client, bucketsMkdir, fileType + separator).then(res => {

        let files = [...res.objects] || [];

        files.map(item => {
            item.time = new Date(item.lastModified).getTime()
            return item;
        })

        let lastFile = _.orderBy(files, ["time"], ['desc']).slice(0, 1)

        let midArr = lastFile[0].name.split(separator)

        let midArrOfTime = midArr[midArr.length - 1].split('.')

        let midTimestamp = midArrOfTime[0]

        let fileName = fileType + separator + midTimestamp + '.' + midArrOfTime[1]

        let localFile = resolve(downloadMkdir + fileName)

        downToLocal(client, lastFile[0].name, localFile).then(_ => {

            returnData.code = 200;
            returnData.msg = `${fileName}文件下载完成`
            response.send(returnData)

        }).catch(err => {
            returnData.code = 0
            returnData.msg = `${fileName}文件下载失败`
            response.send(returnData)
        })
    }).catch(err => {
        returnData.code = 0;
        returnData.msg = "读取OSS服务错误！"
        response.send(returnData)
    })


})
//测试数据
app.post('/getBusinessDataOfTest', (request, response) => {
    let returnData = {
        code: 0,
        msg: "",
        data: []
    }
    let downloadDir = resolve(downloadMkdir)
    readDir(downloadDir, '出差').then(files => {
        let midFiles = []
        files.map(item => {
            let obj = {}
            let midArrOne = item.split(separator)
            let midArrTwo = midArrOne[1].split(".");
            obj.name = item;
            obj.prefix = midArrOne[0]
            obj.time = Number(midArrTwo[0])
            midFiles.push(obj)
        })

        let a = _.orderBy(_.filter(midFiles, {"prefix": "business"}), ["time"], ['desc']).slice(0, 1)
        let objectName = a[0].name;

        if (objectName.length === 0) {
            returnData = {
                code: 0,
                msg: `请先上传出差数据！`,
            }
            response.send(returnData)
        }

        let useFile = resolve(downloadMkdir) + objectName       // userFile : download中最新的文件

        getLocalFileOfTest(useFile, businessFieldsOfTest).then(csvData => {                    // 读取CSV文件内容

            //孵化出差数据和机组数据
            let {richBusinessData} = richData(csvData)

            businessDataOfTest = statisticsBusinessData(richBusinessData)             //孵化后结果

            returnData = {
                code: 200,
                msg: "success",
                data: businessDataOfTest
            }

            response.send(returnData)
        })


    }).catch(err => {

        returnData = {
            code: 0,
            msg: err,
            data: []
        }
        response.send(returnData)
    })
})

app.post('/getHostDataOfTest', (request, response) => {
    let returnData = {
        code: 0,
        msg: "",
        data: []
    }
    let downloadDir = resolve(downloadMkdir)
    readDir(downloadDir, '出差').then(files => {

        let midFiles = []

        files.map(item => {
            let obj = {}
            let midArrOne = item.split(separator)
            let midArrTwo = midArrOne[1].split(".");
            obj.name = item;
            obj.prefix = midArrOne[0]
            obj.time = Number(midArrTwo[0])
            midFiles.push(obj)
        })
        let a = _.orderBy(_.filter(midFiles, {"prefix": "business"}), ["time"], ['desc']).slice(0, 1)
        let objectName = a[0].name;

        if (objectName.length === 0) {
            returnData = {
                code: 0,
                msg: `请先上传出差数据！`,
            }
            response.send(returnData)
        }

        let useFile = resolve(downloadMkdir) + objectName       // userFile : download中最新的文件

        getLocalFileOfTest(useFile, businessFieldsOfTest).then(csvData => {                    // 读取CSV文件内容

            //孵化出差数据和机组数据
            let {richHostData} = richData(csvData)

            hostDataOfTest = statisticsHostData(richHostData)                   // 孵化后结果

            returnData = {
                code: 200,
                msg: "success",
                data: hostDataOfTest
            }

            response.send(returnData)
        })


    }).catch(err => {

        returnData = {
            code: 0,
            msg: err,
            data: []
        }
        response.send(returnData)
    })
})

app.post('/getVideoDataOfTest', (request, response) => {
    let returnData = {
        code: 0,
        msg: "",
        data: {}
    }
    let downloadDir = resolve(downloadMkdir)
    readDir(downloadDir, '视频').then(files => {

        let midFiles = []

        files.map(item => {
            let obj = {}
            let midArrOne = item.split(separator)
            let midArrTwo = midArrOne[1].split(".");
            obj.name = item;
            obj.prefix = midArrOne[0]
            obj.time = Number(midArrTwo[0])
            midFiles.push(obj)
        })
        let a = _.orderBy(_.filter(midFiles, {"prefix": "video"}), ["time"], ['desc']).slice(0, 1)
        let objectName = "/" + a[0].name;

        if (objectName.length === 0) {
            returnData = {
                code: 0,
                msg: `请先上传出差数据！`,
            }
            response.send(returnData)
        }
        returnData = {
            code: 200,
            msg: "上传成功"
        }
        returnData.data = {
            url: objectName,
        }
        response.send(returnData)

    }).catch(err => {
        returnData = {
            code: 0,
            msg: err,
            data: []
        }
        response.send(returnData)
    })
})
// 处理位置接口
app.post('/getHostNameList', (request, response) => {

    let returnData = {
        code: 0,
        msg: "",
        data: []
    }

    let downloadDir = resolve(downloadMkdir)
    readDir(downloadDir).then(files => {

        let objectName = filterDir(files, 'test')

        if (objectName.length === 0) {
            returnData = {
                code: 0,
                msg: `请先上传出差数据！`,
            }
            response.send(returnData)
        }

        let useFile = resolve(downloadMkdir) + objectName       // userFile : download中最新的文件

        getLocalFileOfTest(useFile, businessFieldsOfTest).then(csvData => {                    // 读取CSV文件内容

            //孵化出差数据和机组数据
            let {richHostData} = richData(csvData)          //机组数据

            let arr = []

            let copyFields = ["项目指令", "用户名称"]
            richHostData.map(item => {

                let arrObj = {}
                for (let innerKey in item) {
                    if (copyFields.includes(innerKey)) {
                        arrObj[innerKey] = item[innerKey]
                    }
                }
                arr.push(arrObj)
            })

            returnData = {
                code: 200,
                msg: "success",
                data: arr
            }

            response.send(returnData)
        })


    }).catch(err => {
        console.log("readDir error");
        returnData = {
            code: 0,
            msg: "readdir error",
            data: []
        }
        response.send(returnData)
    })
})


app.listen("3000", (err) => {
    if (!err) {
        console.log("Server is running!");
    }
})


// 文档说明：
// 1.文件上传到OSS上的命名格式    business-文件名-timestamp.xlxs     video-文件名-timestamp.mp4       使用filterDir函数实现最新文件的过滤
// 2.文件下载到本地（download文件夹中）的命名格式： business-timestamp.xlxs    video-timestamp.mp4      使用filterDir函数实现最新文件的过滤
