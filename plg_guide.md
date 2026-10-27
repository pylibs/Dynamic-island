# Dynamic Island: New Features Guide

This guide covers three key features: Plugins, Screen Recording, and the Microphone/Camera privacy indicators. The plugin section is the most extensive since you can write your own custom extensions.

---

## 1. Plugins

### Where Plugins Are Located

Plugin directory: `C:\Users\<user>\.dynamic_island\plugins`

You can open this folder in two ways:
* Go to the **"Plugins"** tab → click **"Folder"**
* Right-click the Island → select **"Plugins Folder"**

On first launch, an example file named `_example.py` is created in this folder. Files starting with `_` or `.` are ignored by the Island. To enable the example, rename it to `example.py` and click **"Reload"** on the Plugins tab (or right-click the Island → **"Reload Plugins"**).

### What Counts as a Plugin

- A single `.py` file placed inside the plugins folder, **or**
- A subfolder containing a `plugin.py` file (useful if your plugin consists of multiple files).

The only mandatory requirement: your plugin must contain a `setup(api)` function. Everything else is optional.

### Minimal Plugin Example

Create a file named `hello.py`:

```python
NAME = "Hello"

def setup(api):
    api.command(["hello", "hi"], "Say Hello",
                lambda: api.flash("Hello!"))
```

Click "Reload", open the launcher with `Alt + Space`, type `hello`, and press Enter. The Island will flash the message "Hello!".

Key points:

- `NAME` is optional. It defines the friendly display name of the plugin shown on the tab. If omitted, the filename is used.
- `setup(api)` is called once upon loading. You register everything your plugin does inside this function.

---

## 2. What `api` Can Do

### `api.compact(fn, idle_only=False)`: Text in the Collapsed Island

`fn` must return a string to display on the Island. If it returns `None` or an empty string, nothing will be shown.

```python
api.compact(lambda: "☕ 3")
api.compact(lambda: "La-la", idle_only=True)   # Only when music is not playing
```

Important:

- The function is called **once every second** and runs on the main thread, so it must be fast (no network calls or reading large files).
- Text is truncated to 16 characters.
- A maximum of 3 plugin compact texts can fit on the Island simultaneously.
- For data fetched from the Internet, use `api.background` to store the result in a variable, and let `compact` simply return that variable.

### `api.command(keys, title, action, glyph="\uE74C", sub=None)`: Launcher Command

Appears in `Alt + Space` when the search query begins with one of the `keys` or matches the `title`.

```python
api.command(["coffee", "cafe"], "Drink Coffee", coffee, glyph="\uE74C")
```

- `keys`: A string or list of keywords. Case-insensitive.
- `action`: A parameterless function executed on Enter.
- `glyph`: Icon code from Segoe Fluent Icons (e.g., `"\uE74C"`).
- `sub`: Gray subtitle text under the title. Defaults to `"Plugin · <plugin_name>"`.
- When the search query is empty, plugin commands are hidden to avoid cluttering the launcher.

### `api.search(fn)`: Custom Launcher Search

`fn(query)` is called on every keypress and returns a list of tuples `(glyph, title, subtitle, action)`.

```python
def search(q):
    if q.startswith("calc "):
        return [("\uE8EF", "Calculate Salary", q[5:], lambda: api.copy(q[5:]))]
    return []

api.search(search)
```

The function must be fast. Plugins collectively display up to 4 search results in the launcher.

### `api.panel(lines_fn=None, buttons=None)`: Card on the "Plugins" Tab

```python
def lines():
    return ["Consumed: 3", "Current time: 14:05"]

api.panel(lines, buttons=[("Reset", reset), ("More", more)])
```

- `lines_fn`: Returns a list of strings. Updated once per second while the tab is open.
- `buttons`: A list of `(label, callback_function)` pairs.

### `api.every(seconds, fn, now=True)`: Recurring Task

```python
api.every(60, check_something)               # Immediately and then every minute
api.every(3600, remind, now=False)           # Only after an hour, not immediately
```

The task runs on the main thread; do not block it.

### `api.after(seconds, fn)`: One-time Task After N Seconds

```python
api.after(10, lambda: api.flash("10 seconds elapsed"))
```

### `api.background(fn, on_done=None)`: Long-running Background Task

Ideal for network calls, disk I/O, and anything that might hang. `fn()` executes in a separate thread, while `on_done(result)` is called back on the main thread.

```python
import requests

def load():
    return requests.get("https://api.example.com/price", timeout=5).json()["usd"]

def show(price):
    state["price"] = price

api.every(300, lambda: api.background(load, show))
api.compact(lambda: f"${state['price']}" if state.get("price") else None)
```

Inside `fn`, you **must not** call `api.flash` or any other methods that modify the user interface. Do those inside `on_done`.

### Handy Utilities

| Method | Description |
|---|---|
| `api.flash(text, ms=1600)` | Shows a temporary notification message on the Island |
| `api.copy(text)` | Copies text to the clipboard |
| `api.open(path_or_url)` | Opens a file, folder, executable, or web URL |
| `api.data_dir` | Dedicated plugin storage directory (`Path`) |
| `api.name` | Display name of the plugin |

### `teardown()`: Cleanup on Reload

If you need to close connections or save state, define a `teardown()` function in your module. It is called during "Reload" and when shutting down plugins. Defining it is optional.

---

## 3. Practical Examples

### Counter on the Island with Buttons

```python
NAME = "Counter"

def setup(api):
    state = {"n": 0}

    api.compact(lambda: f"#{state['n']}" if state["n"] else None)

    def add():
        state["n"] += 1

    def reset():
        state["n"] = 0

    api.panel(lambda: [f"Value: {state['n']}"],
              buttons=[("+1", add), ("Reset", reset)])
```

### Currency Rate in the Collapsed Island

```python
import requests

NAME = "Exchange Rate"

def setup(api):
    state = {"rate": None}

    def fetch():
        r = requests.get("https://api.frankfurter.app/latest",
                         params={"from": "USD", "to": "EUR"}, timeout=6)
        return r.json()["rates"]["EUR"]

    def got(rate):
        state["rate"] = rate

    api.every(900, lambda: api.background(fetch, got))
    api.compact(lambda: f"€{state['rate']:.2f}" if state["rate"] else None)
    api.panel(lambda: [f"1 USD = {state['rate']} EUR" if state["rate"] else "Loading…"])
```

### Save Note to File via Command

```python
NAME = "Notes"

def setup(api):
    def search(q):
        if q.lower().startswith(("n ", "note ")):
            text = q.split(" ", 1)[1]

            def save():
                with open(api.data_dir / "notes.txt", "a", encoding="utf-8") as f:
                    f.write(text + "\n")
                api.flash("Note saved")

            return [("\uE70B", f"Write: {text}", "Notes Plugin", save)]
        return []

    api.search(search)
```

In the launcher, type `note buy bread` and press Enter.

---

## 4. Troubleshooting Plugins

- **Plugin does not appear on the tab.** Check that the file is in the correct directory, its name does not start with `_` or `.`, and its extension is `.py`. Then click "Reload".
- **Red text "⚠ …" on card.** This displays the error message. Detailed tracebacks are written to `C:\Users\<user>\.dynamic_island\plugins.log`.
- **"Plugin has no setup(api) function".** Add `def setup(api):` to your file.
- **Single plugin crashes.** Other plugins continue running normally. Fix the file and click "Reload".
- **Island starts lagging.** Slow operations (network, disk I/O) are likely running inside `compact`, `search`, or `every`. Move them into `api.background`.

Important: Plugins are standard Python code with full system access. Do not install plugins from untrusted sources.

---

## 5. Screen Recording

**Preparation.** Requires `ffmpeg.exe`. Place it in one of the following locations:

- Next to `main.py`;
- In the `C:\Users\<user>\.dynamic_island\` folder;
- Or add `ffmpeg` to your system `PATH`.

**How to start recording:**

- Click the ● button on the main tab (when the Island is expanded);
- Right-click the Island → "Start Screen Recording";
- Open `Alt + Space` launcher and run the `rec` command.

While recording, a blinking red dot and timer appear on the collapsed Island, e.g., `● 01:23`. Stop recording using the same methods: the ■ button, context menu item, or launcher command.

Videos are saved to `C:\Users\<user>\Videos\Island\` in MP4 format (open this directory via right-click → "Recordings Folder").

**Limitations:**

- Video is recorded without audio;
- The Island itself is visible in the video;
- Captures the entire primary display screen.

---

## 6. Microphone and Camera Indicators

No setup required. It works automatically:

- An **orange microphone icon** is displayed when any application is actively capturing audio;
- A **green camera icon** is displayed when any application is using the webcam;
- Hovering over the Island displays a tooltip indicating which application is using the device.

Data is retrieved from Windows Privacy Settings (Windows 10 version 1903+ and Windows 11). If an app crashes unexpectedly, the icon might occasionally stay active until that app's process completely terminates.