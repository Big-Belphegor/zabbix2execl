import requests,json,time,xlwt

header = {
    "Content-Type": "application/json"
}

st = "%s 00:00:00" % (time.strftime('%Y-%m-%d',time.localtime()))
et = "%s 17:30:00" % (time.strftime('%Y-%m-%d',time.localtime()))

class GetZabbix:
    def __init__(self):
        self.username = "Admin"
        self.password = "zabbix"
        self.url = "http://172.16.10.3/zabbix/api_jsonrpc.php"
        self.token = self.getToken()
        self.st = self.timeCovert(st)       # 起始时间，用于定义History取值的时间范围
        self.et = self.timeCovert(et)       # 结束时间

    def timeCovert(self,stringtime):
        '''转换时间格式为Zabbix API可读的数据格式'''
        timeArray = time.strptime(stringtime, "%Y-%m-%d %H:%M:%S")
        timeStamp = int(time.mktime(timeArray))
        return timeStamp

    def getToken(self):
        '''调用Zabbix API所需的Token'''
        data = {
            "jsonrpc": "2.0",
            "method": "user.login",
            "params": {
                "user": self.username,
                "password": self.password
            },
            "id": 0
        }
        ret = requests.post(self.url,headers=header,data=json.dumps(data))
        token = json.loads(ret.content).get("result")
        return token

    def getHosts(self):
        '''返回Zabbix监控主机的hostid、host、name、group等基础信息'''
        data = {
            "jsonrpc": "2.0",
            "method": "host.get",
            "params": {
                "output": [
                    "hostid",
                    "host",
                    "name"
                ],
                "selectInterfaces": [
                    "interfaceid",
                    "ip"
                ],
                "selectGroups": [
                    "group"
                ],
            },
            "id": 2,
            "auth": self.token
        }
        ret = requests.post(url=self.url, headers=header, data=json.dumps(data))
        request = json.loads(ret.text)
        return request['result']

    def getHostGroups(self):
        '''返回Zabbix Server所有Group信息'''
        data = {
            "jsonrpc": "2.0",
            "method": "hostgroup.get",
            "params": {
                "output": "extend",
                "sortfield": [
                    "groupid",
                    "name"
                ],
            },
            "auth": self.token,
            "id": 1
        }
        ret = requests.post(url=self.url, headers=header, data=json.dumps(data))
        request = json.loads(ret.text)
        return request['result']

    def getItem(self,ip,itemkey):
        '''获取ItemID'''
        data = {
            "jsonrpc": "2.0",
            "method": "item.get",
            "params": {
                "host": ip,
                "search": {
                    "key_": itemkey
                },
            },
            "auth": self.token,
            "id": 1
        }
        ret = requests.post(self.url,headers=header,data=json.dumps(data))
        request = json.loads(ret.content).get("result")
        return request[0]['itemid']

    def getHistoryT0(self,itemid):
        '''通过Item ID匹配获取self.st至self.et期间Zabbix的历史数据，此处数据为浮点类型数据'''
        data = {
            "jsonrpc": "2.0",
            "method": "history.get",
            "params": {
                "output": "extend",
                "itemids": itemid,
                "history": 0,               # 浮点型数据
                "time_from": self.st,
                "time_till": self.et,
            },
            "id": 2,
            "auth": self.token,
        }
        ret = requests.post(self.url,headers=header,data=json.dumps(data))
        data = json.loads(ret.content)
        return data['result']

    def getHistoryT3(self,itemid):
        '''通过Item ID匹配获取self.st至self.et期间Zabbix的历史数据，此处数据为整数类型数据'''
        data = {
            "jsonrpc": "2.0",
            "method": "history.get",
            "params": {
                "output": "extend",
                "itemids": itemid,
                "history": 3,               # 数字类型数据
                "time_from": self.st,
                "time_till": self.et,
            },
            "id": 2,
            "auth": self.token,
        }
        ret = requests.post(self.url,headers=header,data=json.dumps(data))
        data = json.loads(ret.content)
        return data['result']

def dumpData(itemkey):
    '''批量导出主机对应Item Key匹配的数据'''
    getzabbix = GetZabbix()
    hostinfo = getzabbix.getHosts()
    data = []
    for i in hostinfo:
        try:
            '''Linux与Windows的Item Key条件不同，故在使用Linux/Windows的条件时不配的系统会报错'''
            itemid = getzabbix.getItem(i['interfaces'][0]['ip'], itemkey)
            values = getzabbix.getHistoryT0(itemid)
            if values:
                tmplist = []                            # 暂存单节点历史区间数据，用于下文运算获取最大值
                for i in values:
                    tmplist.append(float(i["value"]))   # 获取每个时间点的value数据至tmplist列表中
                max_value = max(tmplist)                # 获取最大值
                value = str(round(max_value, 2))        # 去除浮点保留两位小数
                data.append(value)                      # 将所有节点数据存储至data列表中
        except IndexError as e:
            pass
    return data

def dumpData_2(ip,itemkey):
    '''指定导出主机对应Item Key匹配的数据，注意：该方法仅对浮点型数据进行最大值获取'''
    getzabbix = GetZabbix()
    data = []
    try:
        '''Linux与Windows的Item Key条件不同，故在使用Linux/Windows的条件时不配的系统会报错'''
        itemid = getzabbix.getItem(ip, itemkey)
        f_values = getzabbix.getHistoryT0(itemid)   # 浮点型数据
        if f_values:
            tmplist = []                            # 暂存单节点历史区间数据，用于下文运算获取最大值
            for i in f_values:
                tmplist.append(float(i["value"]))   # 获取每个时间点的value数据至tmplist列表中
            max_value = max(tmplist)                # 获取最大值
            value = str(round(max_value, 2))        # 去除浮点保留两位小数
            data.append(value)                      # 将所有节点数据存储至data列表中
    except IndexError as e:
        pass
    return data

def dumpData_3(ip,itemkey):
    '''指定导出主机对应Item Key匹配的数据，注意：该方法仅对数字型数据进行最大值获取'''
    getzabbix = GetZabbix()
    data = []
    try:
        itemid = getzabbix.getItem(ip, itemkey)
        i_values = getzabbix.getHistoryT3(itemid)       # 整数型数据
        if i_values:
            tmplist = []  # 暂存单节点历史区间数据，用于下文运算获取最大值
            for i in i_values:
                tmplist.append(int(i["value"]) / 1024 / 1024 / 1024)  # 转换为GB
            max_value = max(tmplist)
            value = str(round(max_value, 2))
            data.append(value)
    except IndexError as e:
        pass
    return data

def getData():
    '''重构Zabbix API获取出来的数据，便于做下文生成Execl表格的源数据'''
    getzabbix = GetZabbix()
    hostinfo = getzabbix.getHosts()
    groupsinfo = getzabbix.getHostGroups()
    sort_hostinfo = []
    all = []

    for i in hostinfo:
        '''过滤需生成报表主机的GroupID'''
        if i['groups'][0]['groupid'] == "2" and i['groups'][1]['groupid'] != "4" and i['groups'][1][
            'groupid'] != "31" and i['groups'][1]['groupid'] != "28":
            sort_hostinfo.append(i)
    num = sorted(sort_hostinfo, key=lambda n: n['groups'][1]['groupid'])  # 按GroupID排序主机列表

    '''自定义字典格式'''
    for x in num:
        for y in groupsinfo:
            ip = x['interfaces'][0]['ip']               # 获取生成报表节点的IP
            group_id = x['groups'][1]['groupid']        # 获取生成报表节点的GroupID
            hostname = x['name']                        # 获取生成报表节点的Hostname
            if group_id in y['groupid']:                # 若当前GroupID存在于需要生成的报表的主机列表中时，执行
                cpu_idle = dumpData_2(ip, "system.cpu.util[,idle]")             # 导出当前节点CPU数据
                memory_total = dumpData_3(ip,"vm.memory.size[total]")
                memory_pavail = dumpData_2(ip, "vm.memory.size[pavailable]")
                Linux_disk_data_pfree = dumpData_2(ip, "vfs.fs.size[/data,pfree]")
                Linux_disk_data_total = dumpData_3(ip, "vfs.fs.size[/data,total]")
                Linux_disk_pfree = dumpData_2(ip, "vfs.fs.size[/,pfree]")
                Linux_disk_total = dumpData_3(ip, "vfs.fs.size[/,total]")
                all.append({                            # 格式化数据
                    "gn": y['name'],
                    "hostname": hostname,
                    "ip": ip,
                    "cpu_idle": cpu_idle,
                    "memory_total": memory_total,
                    "memory_pavail": memory_pavail,
                    "Linux_disk_data_pfree": Linux_disk_data_pfree,
                    "Linux_disk_data_total": Linux_disk_data_total,
                    "Linux_disk_pfree": Linux_disk_pfree,
                    "Linux_disk_total": Linux_disk_total,
                })
    return all

def setStyle(name,height,bold=False):
    '''设定Excel表格样式'''
    style = xlwt.XFStyle()
    font = xlwt.Font()
    # alignment = xlwt.Alignment()
    # border = xlwt.Borders()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    # alignment.horz = xlwt.Alignment.HORZ_CENTER
    # border.left_coloure = 0x40b
    # border.right = xlwt.Borders.THIN
    # border.top_colour = 0x40b
    # border.bottom = xlwt.Borders.THIN
    return style

def write_data_to_execl():
    '''生成Execl表格'''
    data = getData()
    f = xlwt.Workbook()                                           # 实例化xlwt
    sheet1 = f.add_sheet('报告', cell_overwrite_ok=True)          # 创建Sheet页
    n = 0                                                         # 下文没写一行数据该变量加一，即行数变量
    row0 = ["序号","项目名称",                                    # 表头内容设定，注意：下文填充需对应
            "服务器名称","IP地址",
            "CPU空闲%",
            "内存总量GB","内存剩余%",
            "磁盘'/data'剩余%","磁盘'/data'总量GB",
            "磁盘'/'剩余%","磁盘'/'总量GB"]

    for i in range(0,len(row0)):                                  # 插入表头
        '''首行表头,格式.write(行,列,数据,字符格式)'''
        sheet1.write(0,i,row0[i],setStyle('Times New Roman',220,True))

    for i in data:
        '''表中具体数据填充'''
        n = n + 1
        sheet1.write(n,0,n,setStyle('Times New Roman',220,True))
        sheet1.write(n,1,i['gn'],setStyle('Times New Roman',220))
        sheet1.write(n,2,i['hostname'],setStyle('Times New Roman',220,True))
        sheet1.write(n,3,i['ip'],setStyle('Times New Roman',220))
        sheet1.write(n,4,i['cpu_idle'],setStyle('Times New Roman',220,True))
        sheet1.write(n,5,i['memory_total'],setStyle('Times New Roman',220))
        sheet1.write(n,6,i['memory_pavail'],setStyle('Times New Roman',220,True))
        sheet1.write(n,7,i['Linux_disk_data_pfree'],setStyle('Times New Roman',220))
        sheet1.write(n,8,i['Linux_disk_data_total'],setStyle('Times New Roman',220,True))
        sheet1.write(n,9,i['Linux_disk_pfree'],setStyle('Times New Roman',220))
        sheet1.write(n,10,i['Linux_disk_total'],setStyle('Times New Roman',220))
    f.save(filename)                                                # 保存文件

if __name__ == '__main__':
    filename = r'D:\File\05、日常巡检\生产服务器日常巡检-%s.xls' % (time.strftime('%Y%m%d',time.localtime()))
    write_data_to_execl()

# ps:
# 默认Python不支持对Execl文件修改，若想二次操作需要将文件赋值为变量，然后与新数据一同重写至新文件，然后进行文件名称覆盖
