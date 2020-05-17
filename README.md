# 海報生成系統
## 前言
結合 Google api、ngrok、Flask 所產生的海報生成系統。
## Google api
##### Google Form
此表單用來讓演講者填寫資料，於 Google 上可免費建立該表單。<br>
[Google Form](https://docs.google.com/forms/d/e/1FAIpQLScVogs4pGZJv7y5Q_bwnO73JyrOL3fYi_MXE0vMuYpJauDi5g/viewform)
##### Google Sheet
表單有回覆內容即可打開查看，此為用於存放演講者資料。<br>
[Google Sheet Reply](https://docs.google.com/spreadsheets/d/1Z_nOqiHwL7PcXMwrGYgt7kfX4yMf4iYtvf-AUi4C2dg/edit#gid=2043491560) <br>
[Google Sheet Schedule](https://docs.google.com/spreadsheets/d/1qKvGQkW78uQr1V3ioQaF-aOFcjCIl2WAFSVPzesz7xY/edit#gid=50425730)
##### Google Apps Script
處理 Google Sheet 的資料以及呼叫 Web Api 的地方，可查看 .gs file。<br>
## ngrok
ngrok即為將localhost對應到https domain的服務。
於 ngrok 官網<https://ngrok.com/download>下載解壓縮之後打開 ngrok.yml 找到 authtoken。

      cd ngrok folder/

      $> ngrok authtoken <YOUR_AUTHTOKEN>

      $> ngrok http 5000 
## Flask
利用Flask簡單的設立一個Web Api。<br><br>
安裝 **Flask**
```python
  pip install Flask
```
create **app.py**
```python
  from flask import Flask

  app = Flask(__name__)

  @app.route('/')
  def run():
    return 'Hello World!'
```
## 使用方式 
修改 **app.py**
```python
@app.route('/', methods=['POST'])
def run():
    #print(request.json) # dict
    poster.show(request.json)
    poster.pptx()
    poster.write(request.json)
    poster.read()
    return 'Successfully!'
```
打開 **app.py**
  
    $> python app.py

打開 **ngrok** <br>

    $> ngrok http 5000
    
執行 **.gs file** <br>

```js
  function main(){
    try{
      postRequest();
      sendEmail(s);
      updateSpreadSheet();
    }catch(err){
      sendEmail(m, err);
    }
  }
```
## 環境配置
於 Windows10、python3.6 環境下執行 <br>

    $> pip install -r requirements.txt

## 輸出
speech.pkl
```json
    {
      "演講者": "測試_演講者",
      "演講題目": "測試_演講題目",
      "演講時間": "測試_演講時間",
      "學歷/公司": "測試_學歷/公司",
      "交通工具": "測試_交通工具",
      "電子郵件": "測試_電子郵件"
    }
```