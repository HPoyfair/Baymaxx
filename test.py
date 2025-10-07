import view_clients as vc

vc._ensure_file()
print("Path: ", vc.DATA_PATH, "EXISTS?", vc.DATA_PATH.exists(), "\n")

c = vc.add_client("poopy")
print("added ", c , "\n")

print("all clients: ", vc.list_clients(), "\n")