CSC		= /cygdrive/c/windows/microsoft.net/framework/v4.0.30319/csc.exe
TARGET	= comment.exe
SRC		= \
	main.cs \

DEPS	=

LIB_VER = 15.0.0.0__71e9bce111e9429c
LIB_DIR = C:\\Windows\\assembly\\GAC_MSIL\\

CSC_FLAGS		= /nologo \
	/target:winexe \
	/reference:"$(LIB_DIR)\\office\\$(LIB_VER)\\office.dll" \
	/reference:"$(LIB_DIR)\\Microsoft.Office.Interop.PowerPoint\\$(LIB_VER)\\Microsoft.Office.Interop.PowerPoint.dll"

DEBUG_FLAGS		= 
RELEASE_FLAGS	= 

.PHONY: debug
debug: CSC_FLAGS+=$(DEBUG_FLAGS)
debug: all

.PHONY: release
release: CSC_FLAGS+=$(RELEASE_FLAGS)
release: all

all: $(TARGET)
$(TARGET): $(SRC) $(DEPS)
	$(CSC) $(CSC_FLAGS) /out:$(TARGET) $(SRC) /utf8output

.PHONY: clean
clean:
	rm $(TARGET)


