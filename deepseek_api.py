from openai import OpenAI

client = OpenAI(
    api_key="sk-593929d5ea2e405396e5c6eabb5d1535",  # 建议用环境变量管理API Key，不要硬编码
    base_url="https://api.deepseek.com"
)

response = client.chat.completions.create(
    model="deepseek-chat",
    messages=[
        {"role": "system", "content": "You are a helpful assistant"},
        {"role": "user", "content": "请模仿庞德《在地铁站》，生成一句诗歌，两行，要求不直接抒情，用一个场景表达一个朦胧、若有所思的情感，中文"
                                    




                                    },
    ],
    stream=True  # 流式模式
)

# 正确方式：迭代流式响应
for chunk in response:
    if chunk.choices[0].delta.content:  # 注意是 delta.content，不是 message.content
        print(chunk.choices[0].delta.content, end="", flush=True)  # 逐块打印