from flask import Flask

app = Flask(__name__)

@app.route("/")
def hello():
    return "Hello World"

@app.route("/test")
def test():
    return "Test works"

if __name__ == "__main__":
    print("Starting test app...")
    print("Routes registered:")
    for rule in app.url_map.iter_rules():
        print(f"  {rule.rule} -> {rule.endpoint}")
    app.run(host="0.0.0.0", port=5001, debug=False)
