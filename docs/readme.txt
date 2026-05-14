# 計數方式

1. 推桿推一次計一次，由PLC計算
2. 面板顯示數字是從PLC讀取的
3. 軟體部分，把app.counter計數、SAMPLE_ID計數、送出PLC信號寫在一起
4. sampleId 就是 input.count。 而 input.count 取自app.counter["stop" + camID]

問題：若在送信號之後，實際推料之前急停，會出現偏差?
OK#1 OK#2 先增加 -> SAMPLE_ID增加 ->（隔較久） 推料數增加
OK#1 OK#2 ， 與SAMPLE_ID不符，有影響嗎？ -> 會出現，有檢測出來，但是沒有圖的情況，不影響整體運行
SAMPLE_ID ， 與推料數不符，有影響嗎？ -> 圖檔會出現"多餘的圖"，沒有實際把料推下去，可能會導致報表與實際不符(+1~2)
目前是採取 "暫停馬上停、當下的OK箱不採計"，但是必定會造成軟體報表跟現實數量不對等

SAMPLE_ID沒有的時候一開始怎麼設的 -> 在 switchbutton 跟OK/NG/NULL一起