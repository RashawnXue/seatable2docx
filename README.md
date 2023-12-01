# seatable2docx

This is a script that can transform data in [seatable](https://seatable.cn/) into docx file.

## Info

- Developed with **python3.6**

## Usage

- Fill the `server_url` and `api_token` in `mian.py`

- Custom your data in docx in method `get_data()`

- Run command
```shell
$ pip3 install -r requirements.txt
$ uvicorn main:app --host 0.0.0.0 --port 8000
```
- Now you can access your file download with command:
```shell
$ curl localhost:8000/download -o file.docx
```

## Tips

- The logic in script is just a demo, please use it with your actual demand.