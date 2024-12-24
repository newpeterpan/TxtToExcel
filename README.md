TxtToExcel.bas是一個Excel巨集，用來將AP資訊的文字檔轉為Excel

文字檔的格式如下：
host AP-205_1.01.01 { hardware ethernet 84:d4:7e:c1:53:fa; fixed-address 10.138.21.121;}
我要從這行文字中提取三個資訊：
Host：AP名稱，例如 AP-205_1.01.01
Hardware ethernet：MAC地址，例如 84:d4:7e:c1:53:fa
fixed-address：IP地址，例如 10.138.21.121

所以這個VBA會讀取這個文字檔，然後把它轉成一個Excel檔。這個Excel檔會有三個欄位，分別是：
hostName：記錄AP名稱
macAddress：記錄MAC位址
ipAddress：記錄IP位址

因為這個文字檔的分行符號是 lf，Excel VBA無法用lf來分行，所以使用 Split 函數來處理。
