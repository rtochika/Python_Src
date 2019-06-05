#---------------------------------------------------
#日付データの0を埋めるクラス：2019/1/1→2019/01/01
#dat10に2019/1/1→2019/01/01（10桁）
#dat07に2019/1/1→2019/01（5桁）
#---------------------------------------------------
class DateFormat():#0を埋めるクラス：2019/1/1→2019/01/01
    def __init__(self, date):
        self.date10=date
        self.date07=date

    def format(self):
        lng=len(self.date10)
        #print("Leng:", lng)
        if lng==8:#2019/1/1→2019/01/01
            self.date10=self.date10.replace("/","/0")
        elif lng==9:
            if self.date10[7:8]=="/":#日付が一桁 → 2019/10/1 →2019/10/01
                self.date10 = self.date10[0:8] + "0" + self.date10[8:9]
            else:#月が一桁 → 2019/1/21 →2019/01/21
                self.date10=self.date10[0:5]+"0"+self.date10[5:9]

        self.date07 =self.date10[0:7]#7桁の年月をセット→2019/05

if __name__ == '__main__':
    dt=DateFormat('2019/12/2')
    dt.format()
    print("dt.f_dat10",dt.date10)
    print("dt.f_dat07",dt.date07)

