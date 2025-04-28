from fastapi import FastAPI, Request

app = FastAPI()

@app.post("/trigger")
async def receive_data(request: Request):
    data = await request.json()
    print("受け取ったデータ:", data)
    return {"status": "success", "received": data}
