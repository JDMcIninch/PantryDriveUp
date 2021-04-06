from .server import app, my_ip_address

app.run(host=my_ip_address())