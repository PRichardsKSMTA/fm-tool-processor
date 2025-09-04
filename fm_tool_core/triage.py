from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.sharepoint.folders.folder import Folder

site_url = "https://ksmcpa.sharepoint.com/sites/KSMTA_CCTN"
username = "powerautomatesvcksmta@ksmcpa.onmicrosoft.com"
password = "!af8A57mD3Ab!%BA%vJoMof"

ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))

# Who am I?
cu = ctx.web.current_user.get().execute_query()
print("CurrentUserLogin:", cu.login_name)

# List all document libraries the account can SEE (BaseTemplate 101)
lists = ctx.web.lists.get().execute_query()
doclibs = [l for l in lists if l.properties.get("BaseTemplate") == 101]
print("DocLibs visible to this account:")
for l in doclibs:
    root = l.root_folder.get().execute_query()
    print("-", l.properties.get("Title"), "|", root.serverRelativeUrl)

# Try to read the expected library and folder explicitly
target_lib_title = "Client Downloads"
target_subfolder = "Pricing Tools"

target_lib = next((l for l in doclibs if l.properties.get("Title") == target_lib_title), None)
if target_lib is None:
    print(f"DocLib '{target_lib_title}' NOT visible (no Read).")
else:
    lib_root = target_lib.root_folder.get().execute_query()
    print("Client Downloads root:", lib_root.serverRelativeUrl)
    # Try direct folder get
    full_path = lib_root.serverRelativeUrl.rstrip("/") + "/" + target_subfolder
    try:
        folder = ctx.web.get_folder_by_server_relative_url(full_path).get().execute_query()
        print("Found target folder:", folder.serverRelativeUrl)
    except Exception as e:
        print("Direct get target folder FAILED:", e)

# Brute-search for any folder named "Pricing Tools" the account can see (two levels deep)
found_paths = []
for l in doclibs:
    root = l.root_folder.get().execute_query()
    try:
        subs = root.folders.get().execute_query()
        for f in subs:
            if f.name == target_subfolder:
                found_paths.append(f.serverRelativeUrl)
            # one more level
            try:
                subs2 = f.folders.get().execute_query()
                for f2 in subs2:
                    if f2.name == target_subfolder:
                        found_paths.append(f2.serverRelativeUrl)
            except Exception:
                pass
    except Exception:
        pass

print("Discovered 'Pricing Tools' folders:", found_paths if found_paths else "None")

# If we found the folder, attempt a tiny add with a unique name (no overwrite)
if found_paths:
    import uuid
    test_path = found_paths[0]
    folder = ctx.web.get_folder_by_server_relative_url(test_path)
    test_name = f"perm-probe-{uuid.uuid4().hex}.txt"
    try:
        folder.upload_file(test_name, b"probe").execute_query()
        print("Add test SUCCEEDED at:", test_path)
    except Exception as e:
        print("Add test FAILED at:", test_path, "->", e)