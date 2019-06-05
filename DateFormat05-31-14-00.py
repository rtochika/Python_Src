#---------------------------------------------------
#日付データの0を埋めるクラス：2019/1/1→2019/01/01
#flg=0:"2019/01/01",flg=1:"2019/01"→dateにセット
#---------------------------------------------------
class DateFormat():#0を埋めるクラス：2019/1/1→2019/01/01
    def __init__(self, date,flg):
        self.date=date
        self.flg=flg

    def format(self):
        lng=len(self.date)
        print("Leng:", lng)
        if lng==8:#2019/1/1→2019/01/01
            self.date=self.date.replace("/","/0")
        elif lng==9:
            if self.date[7:8]=="/":#日付が一桁 → 2019/10/1 →2019/10/01
                self.date = self.date[0:8] + "0" + self.date[8:9]
            else:#月が一桁 → 2019/1/21 →2019/01/21
                self.date=self.date[0:5]+"0"+self.date[5:9]

        if self.flg==1:
            self.date =self.date[0:7]
        #print("date_rtn", self.date)

if __name__ == '__main__':
    dt=DateFormat('2019/12/2',1)
    dt.format()
    print("dt.f_dat",dt.date)
    print("dt.f_da2",dt.date)

