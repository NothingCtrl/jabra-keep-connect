# Jabra Wireless Headsets Keep Connect (or Keep Alive)

**Windows OS** only!!!

A simple app to schedule the playing of a sound that prevents a Jabra device from going into standby mode (disconnecting).

### Build app

```bash
# simple build without any resources
pyinstaller --noconsole --onefile main.py
```

```
# build with resources
pyinstaller --onefile --noconsole --name jabra-keep-connect --icon resources\icon.ico --add-data "resources\\icon.ico;resources" main.py
```
