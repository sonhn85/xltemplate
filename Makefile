# Toolchain; override on the command line when needed:
#   make CC=clang CXX=clang++
CC ?= gcc
CXX ?= g++

EXCEL_SDK_PATH ?= Excel2013XLLSDK
ADDIN_VERSION ?= dev

SRC_DIR := src
INC_DIR := include
BUILD_DIR := build
DIST_DIR := dist
XLL_OUT := $(DIST_DIR)/xlTemplate.xll

EXCEL_SDK_SRC_PATH := $(EXCEL_SDK_PATH)/SRC
EXCEL_SDK_INC_PATH := $(EXCEL_SDK_PATH)/INCLUDE
FRAMEWRK_PATH := $(EXCEL_SDK_PATH)/SAMPLES/FRAMEWRK
FRAMEWRK_SRC_PATH := $(FRAMEWRK_PATH)
FRAMEWRK_INC_PATH := $(FRAMEWRK_PATH)
UTHASH_INC_PATH := lib/uthash
TINYTEMPLATE_PATH := lib/tinytemplate

CPPFLAGS := \
	-I$(INC_DIR) \
	-I$(SRC_DIR) \
	-I$(EXCEL_SDK_INC_PATH) \
	-I$(FRAMEWRK_INC_PATH) \
	-I$(UTHASH_INC_PATH) \
	-I$(TINYTEMPLATE_PATH)

CFLAGS := -O2 -DADDIN_VERSION=\"$(ADDIN_VERSION)\"
LDFLAGS := -shared
LDLIBS := -lstdc++

OBJECTS := \
	$(BUILD_DIR)/memorypool.o \
	$(BUILD_DIR)/memorymanager.o \
	$(BUILD_DIR)/framewrk.o \
	$(BUILD_DIR)/excel4workaround.o \
	$(BUILD_DIR)/helper.o \
	$(BUILD_DIR)/tinytemplate.o \
	$(BUILD_DIR)/xltemplate.o

.DEFAULT_GOAL := help

.PHONY: all xll clean help

all: xll

$(BUILD_DIR) $(DIST_DIR):
	mkdir -p $@

$(BUILD_DIR)/memorypool.o: $(FRAMEWRK_SRC_PATH)/memorypool.cpp | $(BUILD_DIR)
	$(CXX) $(CFLAGS) -c -o $@ -I$(FRAMEWRK_INC_PATH) $<

$(BUILD_DIR)/memorymanager.o: $(FRAMEWRK_SRC_PATH)/memorymanager.cpp | $(BUILD_DIR)
	$(CXX) $(CFLAGS) -c -o $@ -I$(FRAMEWRK_INC_PATH) $<

$(BUILD_DIR)/framewrk.o: $(FRAMEWRK_SRC_PATH)/framewrk.c | $(BUILD_DIR)
	$(CC) $(CFLAGS) -c -o $@ \
		-I$(FRAMEWRK_INC_PATH) \
		-I$(EXCEL_SDK_INC_PATH) \
		-I$(EXCEL_SDK_SRC_PATH) \
		$<

$(BUILD_DIR)/excel4workaround.o: $(SRC_DIR)/excel4workaround.c | $(BUILD_DIR)
	$(CC) $(CFLAGS) $(CPPFLAGS) -c -o $@ $<

$(BUILD_DIR)/helper.o: $(SRC_DIR)/helper.c | $(BUILD_DIR)
	$(CC) $(CFLAGS) $(CPPFLAGS) -c -o $@ $<

$(BUILD_DIR)/tinytemplate.o: $(TINYTEMPLATE_PATH)/tinytemplate.c | $(BUILD_DIR)
	$(CC) $(CFLAGS) $(CPPFLAGS) -c -o $@ $<

$(BUILD_DIR)/xltemplate.o: $(SRC_DIR)/xltemplate.c | $(BUILD_DIR)
	$(CC) $(CFLAGS) $(CPPFLAGS) -c -o $@ $<

$(XLL_OUT): $(OBJECTS) | $(DIST_DIR)
	$(CC) $(LDFLAGS) -o $@ $^ $(LDLIBS)

xll: $(XLL_OUT)

clean:
	rm -rf $(BUILD_DIR) $(XLL_OUT)

help:
	@echo "xlTemplate Add-in Build"
	@echo ""
	@echo "Targets:"
	@echo "  all     Build the add-in"
	@echo "  xll     Build $(XLL_OUT)"
	@echo "  clean   Remove generated files"
	@echo "  help    Show this help message"
	@echo ""
	@echo "Configuration:"
	@echo "  ADDIN_VERSION (default: dev)"
	@echo "  EXCEL_SDK_PATH"
	@echo "  CC"
	@echo "  CXX"
