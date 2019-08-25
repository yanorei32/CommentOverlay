# Comment overlay

This software displays comment on the screen and controls PowerPoint.

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

