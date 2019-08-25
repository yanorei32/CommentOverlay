# Comment overlay

This software displays comment on the screen and controls PowerPoint.

<img src=https://user-images.githubusercontent.com/11992915/63655100-56791400-c7be-11e9-95f1-852489bfdaf1.png width=320px>

## How to build

* Windows 10
* Office
* Cygwin / GNU Make

```bash
git clone https://github.com/yanorei32/CommentOverlay
cd CommentOverlay
make
```

## Other requirement

Run this command in command prompt (Administrator)

```
netsh http add urlacl url=http://+:6928/ user=USERNAME
```

## How to use

Run application

```bash
cd CommentOverlay
./launch-sample.sh [fontsize] [framerate]
```

GUI Key bind

```
q: Quit
any key (without q): start presentation
n: Next
p: Previous
```

## Coding Rules
* Tab 4 indent / Unix style new line (LF).

