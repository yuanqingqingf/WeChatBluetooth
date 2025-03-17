// pages/excel/excel.js
// 引入依赖库
// 引入依赖库
const XLSX = require("../../utils/xlsx.mini.min.js");
import { Base } from '../../utils/base.js';
import * as echarts from '../../ec-canvas/echarts';
const ecBLE = require('../../utils/ecBLE.js')
const base = new Base();
const getHead =  '0x68,0x73,0x63,0x6D,0x64,0x20'; // 设备给手机发-包头
const giveHead = '0x68,0x73,0x63,0x6D,0x66,0x22'; // 手机给设备发-包头
const tail='0x0d,0x0A';// 包尾
const wenDuFaValue = '0x70' // 温度阈值
const danWei = '0x71' // 单位
const mingCheng = '0x72'//设备名称
const baoJingOff='0x73' // 是否开启报警开关
const deleteData='0x75' // 删除历史数据
const getData='0x76'//查询历史数据
// 35 - 45   
let chartLine;
function getOption(xData, data_cur, data_his) {
  var option = {
    grid: {
      top: '20rpx',
      left: '2%',
      right: '5%',
      bottom: '0',
      containLabel: true
    },
    tooltip: {
      show: true,
      trigger: 'axis',
    },
    xAxis: {
      type: 'category',
      boundaryGap: false,
      data: xData,
      axisLine: {
        lineStyle: {
          color: '#999',
        }
      }
    },
    yAxis: {
      type: 'value',
      min: 'dataMin', // 最小值取数据中的最小值
      max: 'dataMax',  // 最大值取数据中的最大值
      axisLine: {
        lineStyle: {
          color: '#999',
          opacity: 0
        }
      },
      splitLine: {
        lineStyle: {
          type: 'dotted',
          color: '#D7D7D7'
        }
      }
    },
    series: [{
      data: data_cur || [],
      type: 'line',
      showSymbol: true,
      symbol: 'circle',
      symbolSize: 7,
      itemStyle: {
        normal: {
          label : {show: true},
          // color: '#65E893', //折点颜色
          color: (params) => {
            // 根据值动态设置颜色
            // if (params.value > 38) {
            //     return '#f56c6c'; // 超过 38 显示红色
            // } else {
                return '#65E893'; // 其他显示绿色
            // }
          },
        }
      },
      lineStyle:{
        color: '#dcdfe6' //折线颜色
      }
    }]
  };
  return option;
}
Page({

  /**
   * 页面的初始数据
   */
  data: {
    hsitoryTime:'',
    chuHistoryTime:'',
    isRead:false,
    echartsWen:{
      rgtime:['05:10','06:10','07:10','08:10','09:10','10:10','11:10'],
      value:[36.5, 37.3, 36, 35.8, 38.3,37.8, 38],
    },
    timerId:null,
    isFirst:true,
    historyShow:false,// 历史记录弹窗
    mainColor:'#65E893',
    wenColor:'#fff',
    zIndex:99999,
    setVisible:false,//设置弹窗visible
    name:'',
    setObj:{
      wenValue:'',
      unit:'',
      isWarn:false,
    },
    nameVisible:false,//设备名称-visible
    xlsxpath:'', // 导出地址
    ecLine: {}, 
    WenInfo:{
      name:'',
      temperature:'',
      unit:',单位',
      warnValue:'',
      dianliang:'',
      warnStatus:false, // 报警开关
    },
    xlsxdata:[
      // {
      //   rgtime:"16:25",
      //   value:36.9
      // },{
      //   rgtime:"17:35",
      //   value:37.1
      // }
    ],
  },
  observers: {
    'WenInfo.temperature, WenInfo.warnValue': function (temperature, warnValue) {
      if(WenInfo.warnStatus){
        this.updateTextStyle(temperature, warnValue);
      }
    },
  },
  // 更新样式
  updateTextStyle(temperature, warnValue) {
    const wenColor = this.getTextStyle(temperature, warnValue);
    this.setData({ wenColor });
  },

  // 根据温度值和警告值返回样式对象
  getTextStyle(temperature, warnValue) {
    return  parseFloat(temperature) >= parseFloat(warnValue) ? '#f31d1d' : '#fff'
    
  },
  // 清除历史温度
  deleteWen(){
    base.modal('是否删除历史温度？',()=>{
      wx.removeStorageSync('historyWen')
    })
  },
  /**
   * 下载exportExcel
   */
exportExcel(){
   let that = this
    //表内容
    // let xlsxdata=this.data.xlsxdata
    let xlsxdata = wx.getStorageSync('historyWen') || []
   if(!xlsxdata.length){
     base.toast('暂无历史温度')
     return
   }
    // 表头
    let title = ['时间','温度值'];
    let sheet = [title]
    // 数据整理
    xlsxdata.forEach(item => {
      sheet.push([item.rgtime,item.value])
    });
    console.log('啥？',sheet)
    // return
  // XLSX插件使用
  var ws = XLSX.utils.aoa_to_sheet(sheet);
  var wb = XLSX.utils.book_new();
  let time =  base.formatTime(new Date())
  // console.log('多少',time)
  XLSX.utils.book_append_sheet(wb, ws, time);
  const fileData = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' }); 
  //保存的本地地址
  const filePath = `${wx.env.USER_DATA_PATH}/${time}.xlsx`; 
  // 写文件
  const fs = wx.getFileSystemManager()
    fs.writeFile({
    filePath: filePath,
    data: fileData,
    encoding: 'base64',
    success(res) {
      that.setData({xlsxpath:filePath});
      const sysInfo = wx.getDeviceInfo()
      // 导出
      if (sysInfo.platform.toLowerCase().indexOf('windows') >= 0) {
        // 电脑PC端导出
        wx.saveFileToDisk({
          filePath: filePath,
          success(res) {
            console.log(res)
          },
          fail(res) {
            console.error(res)
            util.tips("导出失败")
          }
        })
      } else{
        // 打开文档,filePath 是写入资源的临时保存路径
        wx.openDocument({
          filePath: filePath,
          showMenu:true,
          success: function (res) {
              console.log('打开文档成功')
          },
          fail: console.error
        })
      }
    },
    fail(res) { 
      console.error('失败',res)
    }
  });
},
timeInterval(){
  if(this.data.isFirst){
    const timerId = setInterval(() => {
      ecBLE.writeBLECharacteristicValue('<CONNECT>', false)
      console.log('给模块发送指令：<CONNECT>')
      this.setData({
        isFirst:false
      })
    }, 110000);
    this.setData({
      timerId: timerId,
    });
  }
},
// 历史数据保存(只保存最近的30000条)
saveHistory(time,value,isHex){
  let timeX =''
  if(isHex){
    let hour = time.substring(0,2) 
    let minute = time.substring(2,4) 
    timeX = `${hour}:${minute}` // 20:54
  }else{
    timeX = time
  }
  let storageList = wx.getStorageSync('historyWen') || []
  storageList.unshift({rgtime:timeX,value})
  let effectiveValue =  storageList.slice(0,30000)
  console.log('保存历史数据',timeX,'和',value)
  wx.setStorageSync('historyWen',effectiveValue)
  if(!isHex){
    this.getStudentDes()
  }
  // console.log('单独',tempDate)
},
hsitorFirstWen(){
  let storageList = wx.getStorageSync('historyWen') || []
  if(storageList.length){
    this.setData({
      ['WenInfo.temperature']: storageList[0].value,
    })
  }else{
    if(!this.data.WenInfo.temperature){
      wx.showLoading({
        title: '请耐心等待1分钟，正在加载温度',
      })
    }
  }
},
// 2s未发送历史数据则证明历史数据为空，历史数据为空，则发送RTON指令
setRton(){
  setTimeout(()=>{
    wx.hideLoading()
    console.log('2s未收到历史消息isRead:',this.data.isRead,'和chuHistoryTime：',this.data.chuHistoryTime)
    if(this.data.isRead && !this.data.chuHistoryTime){
      this.getStudentDes() // 历史温度发送结束，更新echarts图
     this.hsitorFirstWen()
      console.log('2s后未收到历史消息直接发送RTON指令')
      ecBLE.writeBLECharacteristicValue('<RTON>', false)
      this.setData({
        isRead:false
      })
    }
  },2000)
},
  /**
   * 生命周期函数--监听页面加载
   */
  onLoad(options) {
    // wx.removeStorageSync('historyWen')
    let that = this
    // 设置默认报警值和报警开关
    // const name = options.name;
    this.setData({
      ecLine:{onInit: function (canvas, width, height, dpr) {
        //初始化echarts元素，绑定到全局变量，方便更改数据
        chartLine = echarts.init(canvas, null, {
          width: width,
          height: height,
          devicePixelRatio: dpr
        });
        canvas.setChart(chartLine);
        // that.getStudentDes()
      }
    },
    ['WenInfo.warnValue']: wx.getStorageSync('warnValue') || 37,
    ['WenInfo.warnStatus']: wx.getStorageSync('warnStatus') || false,
    // ['WenInfo.name']:name
  })
  if(this.data.WenInfo.warnStatus){
    this.updateTextStyle(this.data.WenInfo.temperature, this.data.WenInfo.warnValue);
  }
    ecBLE.setChineseType(ecBLE.ECBLEChineseTypeGBK)
    /**
     * 这一组为芯片内部存储: 发送INT；收到INT xxx M，发送<TIME时间>；收到TIME OK，发送<Start>；
     * 收到ACK之后，发送INT；收到INT xxx M，发送<TIME时间>；收到TIME OK，发送<Start>；收到Power on，发送RTON
     */
    // 监听蓝牙变化
    ecBLE.onBLECharacteristicValueChange((curStr, strHex) => {
      let str = curStr.substr(2)
      console.log('模块回复:',curStr,'去除前两字符后：',str)
      if(str.startsWith('<PWR')){ // 电量
        let match = str.match(/\d+/); //str=<PWR=2819>
        let text = 0
        console.log('进入电量计算公式',match[0])
        if(match){
          text =parseInt(-320 + 200*((Number(match[0])/4095)*3))
        }
        that.setData({
          ['WenInfo.dianliang']: text+'%',
         })
      }else if(str=='<ACK>'){
        setTimeout(()=>{
          if(this.data.isFirst){
            ecBLE.writeBLECharacteristicValue('<INT=0X01>', false)
            console.log('给模块发送指令：<INT=0X01>设置时间间隔为1分钟')
          }
        },200)
        if(this.data.isFirst){
          wx.showLoading({
            title: '建立通信连接',
          })
        }
      }else if(str.startsWith('INT')){
        setTimeout(()=>{
          let tempDate= Number(base.formatTime(new Date(),false,true))
          console.log('十进制时间：',tempDate)
          let sexDate = tempDate.toString(16).toUpperCase()
          ecBLE.writeBLECharacteristicValue(`<TIME=0x${sexDate}>`, false)
          console.log(`给模块发送指令：<TIME=0x${sexDate}>时间指令`)
        },200)
      }else if(str=='TIME OK'){
        setTimeout(()=>{
          ecBLE.writeBLECharacteristicValue('<START>', false)
          console.log('给模块发送指令：<START>')
        },200) 
      }else if(str=='Power on'){ // 表示正常（int=>time=>RTON）
        setTimeout(()=>{
          this.setData({
            isRead:true,
            hsitoryTime:'',
            chuHistoryTime:''
          })
          ecBLE.writeBLECharacteristicValue('<READ>', false)
          console.log('给模块发送指令：<READ>')
          this.setRton()
        },100) 
        wx.showLoading({
          title: '温度加载中',
        })
        this.timeInterval()
      } else if(str.startsWith('Devicee=')){ // 设备名称
        let temp = str.substr(8)
        this.setData({
          ['WenInfo.name']:temp
        })
      }else if( (str.includes('BAT:') && str.includes('TP:')) || (str.includes('Bat:') && str.includes('Tp:'))){ // 电量BAT和温度TP（BAT:2853,TP:18977）
       wx.hideLoading()
       let lowStr = str.toLowerCase()
      let dian = Number(lowStr.match(/bat:([^,]+)/)[1]) 
      let tp = Number(lowStr.match(/tp:([^,]+)/)[1])
      let a = tp*2.048/32768
      if(tp!=0){ // 温度不为0，再保存
        let temp =  ((1/(1/298.15+(1/3950)*Math.log(a/(2.5-a))))-273.15).toFixed(3)
        this.setData({
          ['WenInfo.temperature']: temp,
         })
         // 将温度保存在本地
         if(str.includes('BAT:')){ // 实时温度
          let tempDate= base.formatTime(new Date(),false,false,true)
          this.saveHistory(tempDate,temp,false)
         }else{ // 历史温度
          let time =  this.data.hsitoryTime
          let curTime = this.data.chuHistoryTime
          if(time === curTime){
            this.saveHistory(time,temp,true)
          }else{
           let xian = base.subtractOneMinute(time)
           console.log('时间计算',time,'和',xian)
           this.saveHistory(xian,temp,true)
            this.setData({
              hsitoryTime:xian,
            })
          }
         }
         console.log('计算的温度',temp)

      }
      if(dian!=0){
        let text =parseInt(-320 + 200*((dian)/4095)*3) // 电量
        this.setData({
          ['WenInfo.dianliang']: text+'%',
         })
         console.log('电量',text)
      }
     
       // <read>历史温度接受成功后，发送<RTON>。历史温度需要存在本地
      } else if(str=='Rtime off'){
        // setTimeout(()=>{
        //   ecBLE.writeBLECharacteristicValue('<READ>', false)
        // },200)
      }else if(str.startsWith('END of')){
        setTimeout(()=>{
          this.hsitorFirstWen()
          ecBLE.writeBLECharacteristicValue('<RTON>', false)
          this.setData({
            isRead:false
          })
          console.log('给模块发送RTON指令')
        },200)
        this.getStudentDes() // 历史温度发送结束，更新echarts图
      }else if(str.includes(',0X') && str.startsWith('0X')){
        // 硬件发送历史温度之前，会先发送时间，eg：0X9531,0X3D2A（去除中间的,0X）
        if(this.setData.isRead){
          let numStr = str.replace(/,0X/g, '') // 16禁止转化为10禁止，2502282054
          let shijinzhi = Number(numStr)+''
          this.setData({
            hsitoryTime:shijinzhi.substr(6,4),
            chuHistoryTime:shijinzhi.substr(6,4), // 保存小时和分钟
          })
          console.log('read历史温度时间，十进制：',numStr)
        }

      }
    //   // 去除前两位
    //  let arr = strHex.split(',')
    //  //['0x17', '0x23', '0x03', '0x75']
    //  let bbb = arr.slice(6,arr.length-2) 
    //   if(bbb[0]==wenDuFaValue){ //温度阈值
    //    let temp =  base.hexDeleteTow(bbb[1]) + '.' + base.hexDeleteTow(bbb[2])
    //    that.setData({
    //     ['WenInfo.warnValue']: temp,
    //    })
    //   }else if(bbb[0]==danWei){ // 单位
    //     let r = ''
    //     if(bbb[1] =='0x01'){
    //       r = '℃'
    //     }else if(bbb[1] =='0x02'){
    //       r = '℉'
    //     }
    //     that.setData({
    //       ['WenInfo.unit']: r,
    //      })
    //   }else if(bbb[0]==mingCheng){ // 名称

    //   }else if(bbb[0]==baoJingOff){ // 报警开关
    //     let r = ''
    //     if(bbb[1] =='0x01'){
    //       r = '开'
    //     }else if(bbb[1] =='0x02'){
    //       r = '关'
    //     }
    //     that.setData({
    //       ['WenInfo.unit']: r,
    //      })
    //   }else if(bbb[0]==deleteData){
    //     // 删除历史数据，啥也不干
    //   }
    //   else if(bbb[0]==getData){ // 查看历史数据

    //   }else{ // 温度-时间（年-月-日-时-分）
    //     // 0x24,0x11,0x24,0x15 ,0x23 后两位是温度
    //     let temp =  base.hexDeleteTow(bbb[bbb.length-2]) + '.' + base.hexDeleteTow(bbb[length-1])
    //    that.setData({
    //     ['WenInfo.temperature']: temp,
    //    })
    //   }
    })
  },
  ceshi(){
     // 获取当前数据
     let echartsWen = this.data.echartsWen;

     echartsWen.rgtime.push('12:10'); // 添加一个新的时间
 
     echartsWen.value.push(37.5); // 添加一个新的百分比
 
     // 更新数据
     this.setData({
       echartsWen: echartsWen,
     });
     this.getStudentDes()
  },
  getStudentDes: function () {
    this.refreshEchar()
  },
  refreshEchar(){
    // const {rgtime,value} = this.data.echartsWen
    let rgtimeList = []
    let value=[]
    const wenList = wx.getStorageSync('historyWen') || []
    console.log('所有历史温度:',wenList)
    // 取前7条即可
    let qian_qi_tiao = wenList.slice(0,30)
    qian_qi_tiao.map(item=>{
      rgtimeList.push(item.rgtime)
      value.push(parseFloat(item.value) )
    })
    var option = getOption(rgtimeList, value);
    chartLine.setOption(option);
    console.log('echarts更新：',wenList)
  },
  /**
   * 打开‘设备名称’弹窗
   */
  openNameDialog(){
    this.setData({
      nameVisible:true,
      name:this.data.WenInfo.name
    })
  },
  /**
   * ’设置‘弹窗’
   */
  openSetDialog(){
    this.setData({
      setObj:{
        wenValue: this.data.WenInfo.warnValue,
        isWarn: this.data.WenInfo.warnStatus,
      } ,
      setVisible:true
    })
  },
  // 设备名称修改
  nameInput(e) {
    this.setData({
        name: e.detail.value
    })
},
nameClose(){
  this.setData({
    nameVisible:false
  })
  setTimeout(()=>{
    this.setData({
      name: '',
    })
  },100)

},
setInfoClose(){
  this.setData({
    setVisible:false,
    
  })
  setTimeout(() => {
    this.setData({
      setObj:{
        wenValue:'',
        unit:'',
        isWarn:false,
      } ,
    })
  }, 100);
},
 // 校验输入内容
 validateInput(value) {
  // 正则表达式：匹配数字和小数点
  const reg = /^\d*\.?\d*$/;
  return reg.test(value);
},
// 温度值
WenValueInput(e) {
  // this.setData({
  //   ['setObj.wenValue']: e.detail.value,
  // })
  const value = e.detail.value; // 获取输入框的值
  const isValid = this.validateInput(value); // 校验输入内容
  if (isValid) {
    // 如果输入合法，更新输入框的值
    this.setData({
      ['setObj.wenValue']: value,
    });
  } else {
    // 如果输入不合法，恢复上一次的合法值
    this.setData({
      ['setObj.wenValue']: this.data.setObj.wenValue,
    });
  }
},
blurWen(){
  this.setData({
    ['setObj.wenValue']: parseFloat(this.data.setObj.wenValue) ,
  });
},
// 温度单位
unitChange(e){
  this.setData({
    ['setObj.unit']: e.detail,
  })
},
// 报警开关
offOnChange(e){
  this.setData({
    ['setObj.isWarn']: e.detail.value,
  })
},
// 设置-弹窗确定
setInfoConfirm(){
  let {isWarn,wenValue} = this.data.setObj
  if(!wenValue){
    base.toast('请输入完整')
    return
  }
  this.setData({
    ['WenInfo.warnValue']: wenValue,
    ['WenInfo.warnStatus']: isWarn
  })
  wx.setStorageSync('warnValue', wenValue)
  wx.setStorageSync('warnStatus', isWarn)
  this.setInfoClose()
  // 要请求硬件之后再改页面上的值吧？
  wx.showToast({
    title: '成功',
    icon: 'success',
    duration: 2000
  })  
},
// 设备名称-确定
nameConfirm(){
  let { name} = this.data
  if(!name){
    base.toast('请输入完整')
    return
  }
  let reg = /^[a-zA-Z0-9_-]{1,10}$/
  if(!reg.test(name)){
    base.toast('只能输入数字、字母、下划线和中划线')
    return
  }
  ecBLE.writeBLECharacteristicValue(`<NAME=${name}>`, false)
  console.log('给模块发送指令：<NAME=${name}>设置设备明证')
  this.setData({
    'WenInfo.name':name
  })
  this.nameClose()
  // 要请求硬件之后再改页面上的值吧？
  wx.showToast({
    title: '成功',
    icon: 'success',
    duration: 2000
  })  
},
// 历史记录弹窗
openHistory(e){
  // setTimeout(() => {
  //   ecBLE.writeBLECharacteristicValue('<RTOFF>', false)
  // }, 200);
  // console.log('给模块发指令<READ>')
  const param = e.currentTarget.dataset.param; // 获取动态数据
  this.setData({
    historyShow:param,
    xlsxdata: wx.getStorageSync('historyWen') || []
  })
  console.log('第三方',wx.getStorageSync('historyWen') || [])
},
  /**
   * 生命周期函数--监听页面初次渲染完成
   */
  onReady() {

  },

  /**
   * 生命周期函数--监听页面显示
   */
  onShow() {

  },

  /**
   * 生命周期函数--监听页面隐藏
   */
  onHide() {

  },

  /**
   * 生命周期函数--监听页面卸载
   */
  onUnload() {
    ecBLE.writeBLECharacteristicValue('<START>', false)
    console.log('给模块发送指令：<START>')
    ecBLE.onBLEConnectionStateChange(() => { })
    ecBLE.onBLECharacteristicValueChange(() => { })
    ecBLE.closeBLEConnection()
    clearInterval(this.data.timerId);
  },

  /**
   * 页面相关事件处理函数--监听用户下拉动作
   */
  onPullDownRefresh() {

  },

  /**
   * 页面上拉触底事件的处理函数
   */
  onReachBottom() {

  },

  /**
   * 用户点击右上角分享
   */
  onShareAppMessage() {

  }
})