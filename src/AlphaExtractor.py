import glob
import os
import shutil
import xml.etree.ElementTree as et
from pathlib import Path
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import font
import urllib.request

EXTRACTABLE_DIRS = ["Defs", "Languages", "Patches"]
CONFIG_VERSION = 3
EXTRACTOR_VERSION = "0.9.0"
WORD_NEWLINE = '\n'
WORD_BACKSLASH = '\\'


# noinspection PyUnusedLocal
class EntryHint(Entry):
    def __init__(self, master=None, hint="", color='grey', **kwargs):
        super().__init__(master, **kwargs)

        self.hint = hint
        self.hint_color = color
        self.default_color = self['fg']

        self.bind("<FocusIn>", self.foc_in)
        self.bind("<FocusOut>", self.foc_out)

        self.put_hint()

    def put_hint(self):
        self.insert(0, self.hint)
        self['fg'] = self.hint_color

    def foc_in(self, *args):
        if self['fg'] == self.hint_color:
            self.delete('0', 'end')
            self['fg'] = self.default_color

    def get(self):
        tmp = super().get()
        if tmp == self.hint:
            return ""
        else:
            return tmp

    def foc_out(self, *args):
        if not self.get():
            self.put_hint()


# noinspection PyUnusedLocal
class CreateToolTip(object):
    def __init__(self, widget, text='widget info'):
        self.waitTime = 500
        self.wrapLength = 180
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.waitTime, self.showtip)

    def unschedule(self):
        tmp = self.id
        self.id = None
        if tmp:
            self.widget.after_cancel(tmp)

    def showtip(self, event=None):
        x = y = 0
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        self.tw = Toplevel(self.widget)
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = Label(self.tw, text=self.text, justify='left',
                      background="#ffffff", relief='solid', borderwidth=1,
                      wraplength=self.wrapLength)
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tw
        self.tw = None
        if tw:
            tw.destroy()


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def loadConfig(fileName='config.dat'):
    with open(fileName, 'r', encoding='UTF8') as fin:
        configs = fin.read().split('\n')

    if CONFIG_VERSION != int(configs[3]):
        raise ValueError()

    gameDir = configs[5]
    modDir = configs[7]
    definedExcludes = configs[9].replace(' ', '').split('/')
    definedIncludes = configs[11].replace(' ', '').split('/')
    exportDir = configs[13]
    exportFile = configs[15]
    isNameTODO = (configs[17] == 'True')
    collisionOption = int(configs[19])

    return gameDir, modDir, definedExcludes, definedIncludes, exportDir, exportFile, isNameTODO, collisionOption


# noinspection PyShadowingNames
def writeConfig(new_gameDir=None, new_modDir=None, new_exportDir=None, new_exportFile=None,
                new_isNameTODO=None, new_collisionOption=None, fileName="config.dat"):
    try:
        gameDir, modDir, definedExcludes, definedIncludes, exportDir, exportFile, isNameTODO, collisionOption = loadConfig()
        if new_gameDir:
            gameDir = new_gameDir
        if new_modDir:
            modDir = new_modDir
        if new_exportDir:
            exportDir = new_exportDir
        if new_exportFile:
            exportFile = new_exportFile
        if str(new_isNameTODO):
            isNameTODO = new_isNameTODO
        if new_collisionOption:
            collisionOption = new_collisionOption
    except (ValueError, FileNotFoundError):
        gameDir = "C:/Program Files (x86)/Steam/steamapps/common/RimWorld"
        modDir = "C:/Program Files (x86)/Steam/steamapps/workshop/content/294100"
        definedExcludes = []
        definedIncludes = "label/description/customLabel/rulesStrings/slateRef/reportString/jobString/verb/labelNoun" \
                          "/gerund/helpText/letterText/labelFemale/labelPlural/text/labelShort/letterLabel" \
                          "/baseInspectLine/beginLetter/rejectInputMessage/deathMessage/beginLetterLabel" \
                          "/recoveryMessage/endMessage/gerundLabel/pawnLabel/onMapInstruction/labelNounPretty" \
                          "/labelForFullStatList/pawnSingular/pawnsPlural/labelSolidTendedWell/labelTendedWell" \
                          "/labelTendedWellInner/destroyedLabel/permanentLabel/skillLabel/graphLabelY/formatString" \
                          "/meatLabel/headerTip/labelMale/ingestCommandString/ingestReportString/useLabel" \
                          "/descriptionFuture/destroyedOutLabel/leaderTitle/textEnemy/textWillArrive/arrivalTextEnemy" \
                          "/letterLabelEnemy/arrivedLetter/stuffAdjective/textFriendly/arrivalTextFriendly" \
                          "/letterLabelFriendly/name/summary/labelMechanoids/approachingReportString" \
                          "/approachOrderString/fuelGizmoLabel/fuelLabel/outOfFuelMessage/labelSocial" \
                          "/calledOffMessage/finishedMessage/discoverLetterLabel/discoverLetterText" \
                          "/discoveredLetterText/discoveredLetterTitle/instantlyPermanentLabel/letter/adjective" \
                          "/labelFemalePlural/successfullyRemovedHediffMessage/offMessage/ingestReportStringEat" \
                          "/renounceTitleMessage/royalFavorLabel/letterTitle".split('/')
        exportDir = "extracted"
        exportFile = "alpha.xml"
        isNameTODO = "True"
        collisionOption = 0

    configs = f"""Alpha's Extractor Configure file
DO NOT EDIT THIS FILE MANUALLY
Config version [int]
{CONFIG_VERSION}
Game directory [string]
{gameDir}
Mod directory [string]
{modDir}
Excluded tags [string, split with '/']
{'/'.join(definedExcludes)}
Included tags [string, split with '/']
{'/'.join(definedIncludes)}
Export directory [string, split with '/']
{exportDir}
Export filename
{exportFile}
Is Name TODO [Boolean, True/False]
{isNameTODO}
Option file collision  [int, 0:Stop, 1:Overwrite, 2:Merge, 3:Refer]
{collisionOption}"""

    with open(fileName, 'w', encoding='UTF8') as fin:
        fin.write(configs)


def loadInit(window):
    frame = Frame(window)
    frame.grid(row=0, column=0, sticky=N + S + E + W)

    for i in range(4):
        Grid.rowconfigure(frame, i, weight=1)
    Grid.columnconfigure(frame, 1, weight=1)

    gameLocation = StringVar(value=gameLoc)
    modsLocation = StringVar(value=modsLoc)

    labelText = f"""림월드가 설치된 폴더와 림월드 창작마당 폴더를 선택해 주십시오.
한 곳의 경로만 선택해도 됩니다.

선택한 림월드 게임 폴더에는 림월드 실행 파일(RimWorldWin64.exe)이 존재해야 합니다.
기본값: [C:/Program Files (x86)/Steam/steamapps/common/RimWorld]

선택한 창작마당 폴더에는 모드 폴더들이 존재해야 합니다.
기본값: [C:/Program Files (x86)/Steam/steamapps/workshop/content/294100]

현재 추출 가능한 폴더는 다음과 같습니다: Defs, Keyed, Strings
모든 종류의 개선안이나, 버그 제보는 항상 환영합니다.

version = {EXTRACTOR_VERSION}"""

    label = Label(frame, text=labelText)
    label.grid(row=0, column=0, columnspan=4, sticky=N + S + E + W)

    def onClick1():
        if text := filedialog.askdirectory():
            gameLocation.set(text)

    def onClick2():
        if text := filedialog.askdirectory():
            modsLocation.set(text)

    label1 = Label(frame, text="림월드 게임")
    label1.grid(row=1, column=0, sticky=E + W, padx=5)
    entry1 = Entry(frame, textvariable=gameLocation)
    entry1.grid(row=1, column=1, columnspan=2, sticky=E + W)
    btn1 = Button(frame, text="폴더 선택", command=onClick1)
    btn1.grid(row=1, column=3, padx=5)

    label2 = Label(frame, text="창작마당")
    label2.grid(row=2, column=0, sticky=E + W, padx=5)
    entry2 = Entry(frame, textvariable=modsLocation)
    entry2.grid(row=2, column=1, columnspan=2, sticky=E + W)
    btn2 = Button(frame, text="폴더 선택", command=onClick2)
    btn2.grid(row=2, column=3, padx=5)

    def onClick2():
        if not (gameLocation.get() or modsLocation.get()):
            messagebox.showerror("경로가 선택되지 않음", "적어도 하나 이상의 경로를 지정해 주십시오.")
            return
        writeConfig(new_gameDir=gameLocation.get(), new_modDir=modsLocation.get())
        global gameLoc, modsLoc
        gameLoc = gameLocation.get()
        modsLoc = modsLocation.get()
        frame.destroy()
        loadSelectMod(window)

    btn = Button(frame, text="추출 가능한 모드 불러오기", command=onClick2)
    btn.grid(row=3, column=0, columnspan=4, sticky=N+S, padx=5, pady=5)


def loadSelectMod(window):
    frame = Frame(window)
    frame.grid(row=0, column=0, sticky=N + S + E + W)
    Grid.rowconfigure(frame, 1, weight=1)
    Grid.columnconfigure(frame, 0, weight=2)
    Grid.columnconfigure(frame, 1, weight=1)

    textLabel1 = Label(frame, text="추출할 모드를 선택하세요")
    textLabel1.grid(row=0, column=0, sticky=N + S + E + W, padx=5, pady=5)

    modListBoxValue = StringVar()
    modListBox = Listbox(frame, selectborderwidth=3, listvariable=modListBoxValue,
                         font=font.Font(family="Courier", size=10))
    modListBox.grid(row=1, column=0, sticky=N + S + E + W)

    searchMod = EntryHint(frame, "[모드 검색]", justify='center')
    searchMod.grid(row=2, column=0, sticky=E + W, padx=5, pady=5)

    textLabel2 = Label(frame, text="추출할 버전 및 폴더들을 선택하세요")
    textLabel2.grid(row=0, column=1, sticky=N + S + E + W, padx=5, pady=5)

    dirList = StringVar()
    dirListBox = Listbox(frame, selectmode='multiple', selectborderwidth=10, listvariable=dirList)
    dirListBox.grid(row=1, column=1, sticky=N + S + E + W)

    extractButton = Button(frame, text="추출하기")
    extractButton.grid(row=2, column=1)

    corePathList = [p.replace('\\', '/') for p in glob.glob(gameLoc + '/Data/*') if os.path.isdir(p)]
    manualModPathList = [p.replace('\\', '/') for p in glob.glob(gameLoc + '/Mods/*') if os.path.isdir(p)]
    workshopModPathList = [p.replace('\\', '/') for p in glob.glob(modsLoc + '/*') if os.path.isdir(p)]
    if not (corePathList + manualModPathList + workshopModPathList):  # If No Mod
        messagebox.showerror("모드 폴더 찾을 수 없음", "선택한 폴더에 어떤 하위 폴더도 존재하지 않습니다.\n프로그램을 종료합니다.")
        exit(0)

    modsNameDict = {}
    sep = ' | '
    for modPath in corePathList:
        modsNameDict[f"   CORE   {sep}{modPath.split('/')[-1]}"] = modPath
    for modPath in manualModPathList + workshopModPathList:
        try:
            name = et.parse(modPath + '/About/About.xml').getroot().find('name').text
            modsNameDict[f"{int(modPath.split('/')[-1]):10d}{sep}{name}"] = modPath
        except (FileNotFoundError, ValueError, AttributeError):
            modsNameDict[f"          {sep}{modPath.split('/')[-1]}"] = modPath
    modsNameDictKeys = list(modsNameDict.keys())
    modsNameDictKeys.sort(key=lambda x: x.split(sep)[1])
    modListBoxValue.set(modsNameDictKeys)

    extractableDirPathList = []
    extractableDirNameList = []

    def onModSelect(evt):
        try:
            modPath = modsNameDict[evt.widget.get(int(evt.widget.curselection()[0]))]
        except IndexError:
            return

        extractableDirPathList.clear()
        extractableDirNameList.clear()
        try:
            modVersionNodeList = list(et.parse(f'{modPath}/LoadFolders.xml').getroot())
            for eachVersionNode in modVersionNodeList:
                for eachLoad in list(eachVersionNode):
                    if not eachLoad.text:
                        continue
                    path = os.path.join(modPath, eachLoad.text) if eachLoad.text != '/' else modPath
                    try:
                        attr = eachLoad.attrib['IfModActive']
                    except KeyError:
                        attr = ""
                    for eachType in os.listdir(path):
                        if eachType in EXTRACTABLE_DIRS:
                            extractableDirPathList.append(os.path.join(path, eachType))
                            if eachType == "Languages":
                                eachType = "Keyed/Strings"
                            name = f"{eachVersionNode.tag} - {eachType}"
                            if attr:
                                name = f"{name} [모드 의존성: {attr}]"
                            extractableDirNameList.append(name)

        except FileNotFoundError:
            for eachLoad in [modPath] + glob.glob(f"{modPath}/*"):
                if os.path.isdir(eachLoad):
                    for eachType in os.listdir(eachLoad):
                        if eachType in EXTRACTABLE_DIRS:
                            extractableDirPathList.append(os.path.join(eachLoad, eachType))
                            ver = eachLoad.split('\\')[-1] if eachLoad != modPath else "default"
                            if eachType == "Languages":
                                eachType = "Keyed/Strings"
                            extractableDirNameList.append(f"{ver} - {eachType}")
        dirList.set(extractableDirNameList)

    modListBox.bind('<<ListboxSelect>>', onModSelect)

    def onExtract():
        global goExtractList

        goExtractList = [extractableDirPathList[idx] for idx in dirListBox.curselection()]
        if not goExtractList:
            messagebox.showerror("폴더를 선택하세요", "추출할 폴더를 하나 이상 선택하세요.")
            return
        frame.destroy()
        loadSelectTags(window)

    extractButton.configure(command=onExtract)

    # noinspection PyUnusedLocal
    def onSearch(evt):
        modSearch = searchMod.get().lower().replace(' ', '')

        if modSearch != "":
            modShow = [modName for modName in modsNameDictKeys if modSearch in modName.lower().replace(' ', '')]
        else:
            modShow = modsNameDictKeys

        modShow.sort(key=lambda x: x.split(sep)[1])
        modListBoxValue.set(modShow)

    searchMod.bind("<KeyRelease>", onSearch)


def parse_recursive(parent, className, tag, lastTag=None, unKnownLiNo=False):
    if list(parent):
        num_list = 0
        for child in list(parent):
            if child.tag == 'li' and not unKnownLiNo:
                yield from parse_recursive(child, className, tag + '.' + str(num_list), lastTag)
                num_list += 1
            else:
                yield from parse_recursive(child, className, tag + '.' + child.tag, child.tag)
    else:
        yield className, lastTag, tag, (parent.text.replace('&', '&amp;').replace('<', '&lt;') if parent.text else "")


def extractDefs(root):
    if 'value' != root.tag and 'Defs' != root.tag:
        raise ValueError("첫 태그가 Defs가 아닙니다. 오류를 발생시킨 모드 이름을 Alpha에게 제보해주세요.")

    for item in list(root):
        try:
            if item.attrib['Abstract'].lower() == 'true':
                continue
        except KeyError:
            pass

        try:
            defName = item.find('defName').text
        except AttributeError:
            continue

        className = item.tag
        if className == 'Def':
            try:
                className = item.attrib['Class']
            except Exception:
                raise ValueError("Def임에도 불구하고 클래스 이름이 없습니다. 오류를 발생시킨 모드 이름과 파일명을 Alpha에게 제보해주세요.")

        yield from parse_recursive(item, className, defName)


def xpathAnalysis(xpath):
    try:
        if '..' in xpath:
            return False, 1
        if '(' in xpath:
            return False, 2
        defNameIndex = -1
        xpath = xpath.split('/')
        for i in range(len(xpath)):
            if '[' in xpath[i]:
                if 'defName' in xpath[i]:
                    defNameIndex = i
                    className, xpath[i] = xpath[i].split('[')
                    xpath[i] = xpath[i].split(']')[0]
                    if '!=' in xpath[i] or xpath[i].count('=') > 1:
                        return False, 3
                    isThereDefName = False
                    defName = None
                    for elem in xpath[i].split('='):
                        if elem.replace(' ', '') == 'defName':
                            isThereDefName = True
                        else:
                            defName = eval(elem)
                    if not isThereDefName or not defName:
                        return False, 4
                    xpath[i] = defName
                elif 'li' in xpath[i]:
                    xpath[i] = xpath[i].split('[')[1].split(']')[0]
                    if '!=' in xpath[i] or xpath[i].count('=') > 1:
                        return False, 5
                    for elem in xpath[i].split('='):
                        if '\"' in elem or '\'' in elem:
                            xpath[i] = eval(elem)
                elif defNameIndex != -1:
                    return False, 6
        if defNameIndex == -1:
            return False, 7
        return className, xpath[defNameIndex:]
    except Exception:
        return False, -1


def analysisOperation(node, modDepend):
    if (operation := node.attrib['Class']) == 'PatchOperationFindMod':
        modDepend.extend([li.text for li in list(node.find('mods'))])
        yield from analysisOperation(node.find('match'), modDepend)
    elif operation == 'PatchOperationSequence':
        for li in list(node.find('operations')):
            yield from analysisOperation(li, modDepend)
    elif operation == 'PatchOperationInsert':
        try:
            xpath = node.find('xpath').text
            className, tagList = xpathAnalysis(xpath)
            if not className:
                print(f'다음 xpath는 {tagList}번 사유로 파싱할 수 없음: {xpath}')
                return
            yield from parse_recursive(node.find('value'), className, '.'.join(tagList[:-1]), lastTag=tagList[-2], unKnownLiNo=True)
        except Exception as e:
            print(e)
            return
    elif operation == 'PatchOperationAdd':
        try:
            xpath = node.find('xpath').text
            if xpath.replace('/', '') == 'Defs':
                yield from extractDefs(node.find('value'))
            className, tagList = xpathAnalysis(xpath)
            if not className:
                print(f'다음 xpath는 {tagList}번 사유로 파싱할 수 없음: {xpath}')
                return
            yield from parse_recursive(node.find('value'), className, '.'.join(tagList), lastTag=tagList[-1])
        except Exception as e:
            print(e)
            return
    elif operation in ['PatchOperationReplace', 'PatchOperationRemove']:
        try:
            xpath = node.find('xpath').text
            print('노드를 변경하는 패치는 추출하지 않음, xpath:', xpath)
        except Exception as e:
            print('노드를 변경하는 패치는 추출하지 않음, xpath 존재하지 않음')
            return
    else:
        return


def extractPatches(root):
    assert 'Patch' == root.tag, "첫 태그가 Patch가 아닙니다. 오류를 발생시킨 모드 이름을 Alpha에게 제보해주세요."

    for item in list(root):
        assert 'Operation' == item.tag, "Patch 하위 태그가 Operation이 아닙니다. 오류를 발생시킨 모드 이름을 Alpha에게 제보해주세요."
        yield from analysisOperation(item, [])


def loadSelectTags(window):
    global goExtractList, excludes, defaults, includes
    global excludeHide, defaultHide, includeHide, dict_class, dict_keyed, list_strings

    frame = Frame(window)
    frame.grid(row=0, column=0, sticky=N + S + E + W)

    Grid.rowconfigure(frame, 0, weight=1)  # label
    Grid.rowconfigure(frame, 1, weight=9)  # list
    Grid.rowconfigure(frame, 2, weight=1)  # moveButton
    Grid.rowconfigure(frame, 3, weight=1)  # search
    Grid.rowconfigure(frame, 4, weight=1)  # extract
    for i in range(3):
        Grid.columnconfigure(frame, i, weight=1)

    dict_tags_text = {}

    for goExtract in goExtractList:
        if goExtract.split('\\')[-1] == 'Defs':
            GoExtractLists = glob.glob(goExtract + "/**/*.xml", recursive=True)
            for path in GoExtractLists:
                try:
                    extracts = extractDefs(et.parse(path).getroot())
                except ValueError as e:
                    messagebox.showerror("에러 발생", str(e) + "\n파일명: " + path)
                    return

                for className, lastTag, tag, text in extracts:
                    try:
                        dict_class[className].append((lastTag, tag, text))
                    except KeyError:
                        dict_class[className] = [(lastTag, tag, text)]
                    if type(text) == list:
                        text = WORD_NEWLINE.join(text)
                    try:
                        dict_tags_text[lastTag].append(text)
                    except KeyError:
                        dict_tags_text[lastTag] = [text]

        elif goExtract.split('\\')[-1] == 'Languages':
            GoExtractLists = glob.glob(goExtract + "\\English\\Keyed\\**\\*.xml", recursive=True)
            for path in GoExtractLists:
                try:
                    nodes = list(et.parse(path).getroot())
                except ValueError as e:
                    messagebox.showerror("에러 발생", str(e) + "\n파일명: " + path)
                    return
                for node in nodes:
                    dict_keyed[node.tag] = node.text if node.text else ""
            list_strings = glob.glob(goExtract + "\\English\\Strings\\**\\*.txt", recursive=True)

        elif goExtract.split('\\')[-1] == 'Patches':
            GoExtractLists = glob.glob(goExtract + "/**/*.xml", recursive=True)
            for path in GoExtractLists:
                try:
                    extracts = extractPatches(et.parse(path).getroot())
                except ValueError as e:
                    messagebox.showerror("에러 발생", str(e) + "\n파일명: " + path)
                    return

                for extract in extracts:
                    if extract:
                        className, lastTag, tag, text = extract
                    else:
                        continue
                    try:
                        dict_class[className].append((lastTag, tag, text))
                    except KeyError:
                        dict_class[className] = [(lastTag, tag, text)]
                    if type(text) == list:
                        text = WORD_NEWLINE.join(text)
                    try:
                        dict_tags_text[lastTag].append(text)
                    except KeyError:
                        dict_tags_text[lastTag] = [text]
        else:
            messagebox.showerror("에러 발생", "Defs, Patches, Keyed, Strings 이외의 폴더는 아직 추출할 수 없습니다.\n자동으로 제외합니다.")

    if dict_tags_text == {}:
        if dict_keyed or list_strings:
            loadSelectExport(window)
        else:
            messagebox.showerror("에러 발생", "추출 가능한 xml가 존재하지 않거나, 찾을 수 없습니다. 프로그램을 종료합니다.")
            window.destroy()
        return

    def showTexts(evt):
        w = evt.widget
        tag = w.get(int(w.curselection()[0]))
        dialog = Toplevel(window)
        dialog.title(tag)
        text = Text(dialog)
        text.insert(1.0, '\n'.join(sorted(list(set(dict_tags_text[tag])))))
        text.configure(state='disabled')
        text.bind("<Escape>", lambda x: dialog.destroy())
        text.bind("<q>", lambda x: moveTag(tag, 0, dialog=dialog))
        text.bind("<w>", lambda x: moveTag(tag, 1, dialog=dialog))
        text.bind("<e>", lambda x: moveTag(tag, 2, dialog=dialog))
        text.focus_set()
        Grid.rowconfigure(dialog, 0, weight=1)
        Grid.columnconfigure(dialog, 0, weight=1)
        text.grid(row=0, column=0, sticky=N + S + E + W)

    excludes = []
    defaults = sorted(dict_tags_text.keys())
    includes = []

    while True:
        candidate = defaults.pop(0)
        try:
            int(candidate)
            excludes.append(candidate)
        except ValueError:
            defaults.insert(0, candidate)
            break

    for _ in range(len(defaults)):
        candidate = defaults.pop(0)
        if candidate in defExcludes:
            excludes.append(candidate)
        elif candidate in defIncludes:
            includes.append(candidate)
        else:
            defaults.append(candidate)

    Label(frame, text="추출 제외 태그").grid(row=0, column=0, sticky=N + S + E + W)
    excludeVar = StringVar(value=excludes)
    excludeList = Listbox(frame, listvariable=excludeVar)
    excludeList.grid(row=1, column=0, sticky=N + S + E + W)
    excludeList.bind("<w>", lambda x: moveTag(excludeList.get(excludeList.curselection()[0]), 1))
    excludeList.bind("<e>", lambda x: moveTag(excludeList.get(excludeList.curselection()[0]), 2))

    Label(frame, text="미분류 태그\n(추출 제외)").grid(row=0, column=1, sticky=N + S + E + W)
    defaultVar = StringVar(value=defaults)
    defaultList = Listbox(frame, listvariable=defaultVar)
    defaultList.grid(row=1, column=1, sticky=N + S + E + W)
    defaultList.bind("<q>", lambda x: moveTag(defaultList.get(defaultList.curselection()[0]), 0))
    defaultList.bind("<e>", lambda x: moveTag(defaultList.get(defaultList.curselection()[0]), 2))

    Label(frame, text="추출 대상 태그").grid(row=0, column=2, sticky=N + S + E + W)
    includeVar = StringVar(value=includes)
    includeList = Listbox(frame, listvariable=includeVar)
    includeList.grid(row=1, column=2, sticky=N + S + E + W)
    includeList.bind("<q>", lambda x: moveTag(includeList.get(includeList.curselection()[0]), 0))
    includeList.bind("<w>", lambda x: moveTag(includeList.get(includeList.curselection()[0]), 1))

    def moveTag(tag, destination, dialog=None):
        if dialog:
            dialog.destroy()
        try:
            toMove = defaults.pop(defaults.index(tag))
            if destination == 1:
                return
            if defaultList.get(END) == toMove:
                defaultList.activate(defaultList.size() - 2)
                defaultList.selection_set(defaultList.size() - 2)
        except ValueError:
            try:
                toMove = excludes.pop(excludes.index(tag))
                if destination == 0:
                    return
                if excludeList.get(END) == toMove:
                    excludeList.activate(excludeList.size() - 2)
                    excludeList.selection_set(excludeList.size() - 2)
            except ValueError:
                try:
                    toMove = includes.pop(includes.index(tag))
                    if destination == 2:
                        return
                    if includeList.get(END) == toMove:
                        includeList.activate(includeList.size() - 2)
                        includeList.selection_set(includeList.size() - 2)
                except ValueError:
                    return

        [excludes, defaults, includes][destination].append(toMove)

        excludeVar.set(sorted(excludes))
        defaultVar.set(sorted(defaults))
        includeVar.set(sorted(includes))

        if destination == 0:
            for i, v in enumerate(excludeList.get(0, END)):
                if v == toMove:
                    excludeList.see(i)
                    break
        elif destination == 0:
            for i, v in enumerate(defaultList.get(0, END)):
                if v == toMove:
                    defaultList.see(i)
                    break
        elif destination == 0:
            for i, v in enumerate(includeList.get(0, END)):
                if v == toMove:
                    includeList.see(i)
                    break

    excludeList.bind('<Double-Button-1>', showTexts)
    defaultList.bind('<Double-Button-1>', showTexts)
    includeList.bind('<Double-Button-1>', showTexts)

    Label(frame, text="[Q]를 입력해 추출 제외").grid(row=2, column=0)
    Label(frame, text="[W]를 입력해 분류 취소").grid(row=2, column=1)
    Label(frame, text="[E]를 입력해 추출 추가").grid(row=2, column=2)

    searchTag = EntryHint(frame, "[태그 필터]")
    searchTag.grid(row=3, column=0, sticky=E + W)
    searchText = EntryHint(frame, "[원본 텍스트 필터]")
    searchText.grid(row=3, column=1, columnspan=2, sticky=E + W)

    dictTextTag = {}
    for tag, texts in dict_tags_text.items():
        for text in texts:
            try:
                dictTextTag[text].append(tag)
            except KeyError:
                dictTextTag[text] = [tag]

    # noinspection PyUnusedLocal
    def onSearch(evt):
        global excludes, defaults, includes, excludeHide, defaultHide, includeHide
        tagSearch = searchTag.get().lower().replace(' ', '')
        tagFlag = (tagSearch != "")
        textSearch = searchText.get().lower().replace(' ', '')
        textFlag = (textSearch != "")

        excludes = excludes + excludeHide
        defaults = defaults + defaultHide
        includes = includes + includeHide

        filteredTags = []
        if textFlag:
            for tags in [dictTextTag[text] for text in dictTextTag.keys() if
                         textSearch in text.lower().replace(' ', '')]:
                filteredTags.extend(tags)
            if tagFlag:
                filteredTags = [tag for tag in list(set(filteredTags)) if tagSearch in tag.lower().replace(' ', '')]
        elif tagFlag:
            filteredTags = [tag for tag in dict_tags_text.keys() if tagSearch in tag.lower().replace(' ', '')]

        if textFlag or tagFlag:
            excludeIntersect = list(set(excludes) & set(filteredTags))
            excludeHide = [tag for tag in excludes if tag not in excludeIntersect]
            excludes = excludeIntersect
            defaultIntersect = list(set(defaults) & set(filteredTags))
            defaultHide = [tag for tag in defaults if tag not in defaultIntersect]
            defaults = defaultIntersect
            includeIntersect = list(set(includes) & set(filteredTags))
            includeHide = [tag for tag in includes if tag not in includeIntersect]
            includes = includeIntersect
        else:
            excludeHide = []
            defaultHide = []
            includeHide = []

        excludeVar.set(sorted(excludes))
        defaultVar.set(sorted(defaults))
        includeVar.set(sorted(includes))

    searchTag.bind("<KeyRelease>", onSearch)
    searchText.bind("<KeyRelease>", onSearch)

    def finishTag():
        frame.destroy()
        loadSelectExport(window)

    Button(frame, text="태그 선택 완료", command=finishTag).grid(row=4, column=1, sticky=N + S + E + W, padx=2, pady=2)

    def loadTagList():
        global excludes, defaults, includes
        if not messagebox.askyesno("태그 분류 불러오기",
                                   "사용자가 사전에 분류했던 태그 분류 작업을 적용합니다.\n" +
                                   "기존 분류 작업의 내용은 보존되지 않으며, 불러온 분류 파일로 덮어씌워집니다. 정말 진행할까요?"):
            return
        fileName = filedialog.askopenfilename(title="불러오기", filetypes=[("태그 분류 작업 파일", "*.tag")])
        if fileName == "":
            return
        try:
            with open(fileName, 'r') as fin:
                reads = fin.read().split('\n')

            customExcludes = reads[1].split('/')
            customIncludes = reads[3].split('/')

            defaults = excludes + defaults + includes
            excludes = []
            includes = []

            for _ in range(len(defaults)):
                candidate = defaults.pop(0)
                if candidate in customExcludes:
                    excludes.append(candidate)
                elif candidate in customIncludes:
                    includes.append(candidate)
                else:
                    defaults.append(candidate)

            excludeVar.set(sorted(excludes))
            defaultVar.set(sorted(defaults))
            includeVar.set(sorted(includes))
            messagebox.showinfo("불러오기 완료", "태그 분류를 성공적으로 불러왔습니다. 기존 작업은 버려졌습니다.")
        except FileNotFoundError:
            messagebox.showerror("파일을 열 수 없습니다.", "파일을 여는 중 오류가 발생했습니다.\n파일이 삭제되었거나 이동했을 수 있습니다.")
            return
        except IndexError:
            messagebox.showerror("분류 파일 불러오기 실패", "분류 파일을 불러오는 데 실패하였습니다.\n분류 파일의 형식이 잘못되었을 수 있습니다.")
            return

    def saveTagList():
        global excludes, includes
        fileName = filedialog.asksaveasfilename(title="저장", filetypes=[("태그 분류 작업 파일", "*.tag")])
        if fileName == "":
            messagebox.showerror("저장 취소", "저장을 취소하였습니다.")
            return
        if fileName[-4:].lower() != '.tag':
            fileName += '.tag'
        writings = '\n'.join(("Excludes tag list, split with [/], spacing ignored, case sensitive", '/'.join(excludes),
                              "Includes tag list, split with [/], spacing ignored, case sensitive", '/'.join(includes)))
        with open(fileName, 'w') as fin:
            fin.write(writings)
        messagebox.showinfo("저장하기 완료", "태그 분류를 성공적으로 저장했습니다.")

    save = Button(frame, command=saveTagList, text="태그 분류 저장하기")
    load = Button(frame, command=loadTagList, text="태그 분류 불러오기")
    save.grid(row=4, column=0)
    load.grid(row=4, column=2)


def loadSelectExport(window):
    global includes

    frame = Frame(window)
    frame.grid(row=0, column=0, sticky=N + S + E + W)

    for i in range(8):
        Grid.rowconfigure(frame, i, weight=1)
    for i in range(1):
        Grid.columnconfigure(frame, i, weight=1)

    varDirName = StringVar(value=exDir)
    varFileName = StringVar(value=exFile)
    varIsNameTODO = BooleanVar(value=isTODO)
    varCollisionOption = IntVar(value=colOption)

    Label(frame, text="결과 폴더 이름 지정").grid(row=0, column=0)
    Entry(frame, textvariable=varDirName).grid(row=1, column=0, sticky=E + W)

    Label(frame, text="결과 파일 이름 지정, [.xml]을 끝에 붙일 것").grid(row=2, column=0)
    Entry(frame, textvariable=varFileName).grid(row=3, column=0, sticky=E + W)

    Label(frame, text="번역해야 할 부분의 텍스트 지정").grid(row=4, column=0)
    row5Frame = Frame(frame)
    row5Frame.grid(row=5, column=0)
    Radiobutton(row5Frame, text="TODO", value=True, variable=varIsNameTODO).grid(row=0, column=0)
    Radiobutton(row5Frame, text="원문", value=False, variable=varIsNameTODO).grid(row=0, column=1)

    Label(frame, text="결과 폴더와 파일명이 일치할 경우 파일의 작업 방법").grid(row=6, column=0)
    row7Frame = Frame(frame)
    row7Frame.grid(row=7, column=0)
    btnTexts = ["중단하기", "덮어쓰기", "병합하기", "참조하기"]
    tooltips = ["파일 충돌이 발생할 경우, 파일 출력을 중단하고 알림을 표시합니다. " +
                "이 경우, 파일 충돌이 발생할 때까지 작성된 파일은 남아 있음에 유의하십시오.",
                "파일 충돌이 발생할 경우, 해당 파일을 내용을 삭제하고 새로 작성합니다. " +
                "이 경우, 기존 파일을 복구할 수 없으므로 사전 백업이 권장됩니다.",
                "파일 충돌이 발생할 경우, 해당 파일에 새로운 태그들을 추가해 병합합니다. " +
                "추출기가 추출하지 않았지만 존재했던 내용은 파일의 하단에 출력됩니다.",
                "파일 충돌이 발생할 경우, 해당 파일의 내용을 새로 작성하되, 기존 파일에 같은 태그가 있을 경우 해당 내용을 보존합니다. " +
                "추출기가 추출하지 않았지만 존재했던 내용은 버려집니다."]
    for i, (text, tooltip) in enumerate(zip(btnTexts, tooltips)):
        tmp = Radiobutton(row7Frame, text=text, value=i, variable=varCollisionOption)
        CreateToolTip(tmp, tooltip)
        tmp.grid(row=0, column=i)

    def exportXml():
        savedList = []
        for className, values in dict_class.items():  # Defs / Patches
            isThereNoIncludes = True
            for lastTag, _, _ in values:  # check existence of include tag
                if lastTag in includes:
                    isThereNoIncludes = False
                    break
            if isThereNoIncludes:
                continue  # pass the class

            filename = varDirName.get() + '/DefInjected/' + className + '/' + varFileName.get()
            Path('/'.join(filename.split('/')[:-1])).mkdir(parents=True, exist_ok=True)

            if varCollisionOption.get() == 0:  # collision -> stop
                try:
                    with open(filename, 'r', encoding='UTF8') as _:
                        messagebox.showerror("파일 충돌 발견됨",
                                             f"다음 폴더의 파일이 존재하여 작업을 중단하였습니다.\n{className}\n\n" +
                                             f"이미 저장된 파일의 폴더 리스트는 아래와 같습니다.\n{WORD_NEWLINE.join(savedList)}")
                        return
                except FileNotFoundError:
                    pass

            alreadyDefinedDict = {}
            if varCollisionOption.get() > 1:  # collision -> merge/refer
                try:
                    for node in list(et.parse(filename).getroot()):
                        if list(node):
                            alreadyDefinedDict[node.tag] = [li.text for li in list(node)]
                        elif node.text != "TODO":
                            alreadyDefinedDict[node.tag] = node.text
                except FileNotFoundError:
                    pass

            writingTextList = []
            for lastTag, tag, text in values:
                if lastTag in includes:
                    if type(text) == list:
                        try:
                            for i, v in enumerate(alreadyDefinedDict[tag]):
                                if alreadyDefinedDict[tag][i] != "TODO":
                                    text[i] = alreadyDefinedDict[tag][i]
                            del alreadyDefinedDict[tag]
                        except TypeError:
                            pass
                        tmp = [f'    <!-- {text_i} -->{WORD_NEWLINE}    <li>TODO</li>' for text_i in text]
                        writingTextList.append(
                            f"  <{tag}>\n{WORD_NEWLINE.join(tmp)}\n  </{tag}>"
                            if varIsNameTODO.get()
                            else f"  <{tag}>\n{WORD_NEWLINE.join([f'    <li>{t}</li>' for t in text])}\n  </{tag}>")
                    else:
                        try:
                            writingTextList.append(f"  <!-- {text} -->\n  <{tag}>{alreadyDefinedDict[tag]}</{tag}>"
                                                   if varIsNameTODO.get()
                                                   else f"  <{tag}>{alreadyDefinedDict[tag]}</{tag}>")
                            del alreadyDefinedDict[tag]
                        except KeyError:
                            writingTextList.append(f"  <!-- {text} -->\n  <{tag}>TODO</{tag}>"
                                                   if varIsNameTODO.get()
                                                   else f"  <{tag}>{text}</{tag}>")

            with open(filename, 'w', encoding='UTF8') as fin:
                fin.write("""<?xml version="1.0" encoding="utf-8"?>\n<LanguageData>\n""")
                fin.write('\n'.join(writingTextList))
                if varCollisionOption.get() == 2 and alreadyDefinedDict:  # collision -> merge
                    fin.write("\n\n  <!-- 알파의 추출기는 추출하지 않았지만 이미 존재했던 노드들 \n\n")
                    for tag, text in alreadyDefinedDict.items():
                        if type(text) == list:
                            fin.write(f"  <{tag}>\n")
                            fin.write('\n'.join([f"    <li>{eachText}</li>" for eachText in text]))
                            fin.write(f"\n  </{tag}>\n")
                        else:
                            fin.write(f"  <{tag}>{text}</{tag}>\n")
                    fin.write("  -->")
                fin.write("\n</LanguageData>")

            savedList.append(className)

        # Keyed
        if len(dict_keyed) > 0:
            filename = varDirName.get() + '/Keyed/' + varFileName.get()
            Path('/'.join(filename.split('/')[:-1])).mkdir(parents=True, exist_ok=True)

            if varCollisionOption.get() == 0:  # collision -> stop
                try:
                    with open(filename, 'r', encoding='UTF8') as _:
                        messagebox.showerror("파일 충돌 발견됨",
                                             f"Keyed 폴더의 파일이 존재하여 작업을 중단하였습니다.\n" +
                                             f"이미 저장된 파일의 폴더 리스트는 아래와 같습니다.\n{WORD_NEWLINE.join(savedList)}")
                        return
                except FileNotFoundError:
                    pass

            alreadyDefinedDict = {}
            if varCollisionOption.get() > 1:  # collision -> merge/refer
                try:
                    for node in list(et.parse(filename).getroot()):
                        if list(node):
                            messagebox.showerror("기존 번역에서 내부 노드 발견됨",
                                                 "본 프로그램은 Keyed 파일의 텍스트에 하위 노드(<>)가 없는 것으로 가정하였습니다. " +
                                                 "해당 노드의 번역은 보존되지 않았을 수 있습니다.")
                            alreadyDefinedDict[node.tag] = [li.text for li in list(node)]
                        elif node.text != "TODO":
                            alreadyDefinedDict[node.tag] = node.text
                except FileNotFoundError:
                    pass

            writingTextList = []
            for tag, text in dict_keyed.items():
                text = text.replace('<', '&lt;')
                try:
                    writingTextList.append(f"  <!-- {text} -->\n  <{tag}>{alreadyDefinedDict[tag]}</{tag}>"
                                           if varIsNameTODO.get()
                                           else f"  <{tag}>{alreadyDefinedDict[tag]}</{tag}>")
                    del alreadyDefinedDict[tag]
                except KeyError:
                    writingTextList.append(f"  <!-- {text} -->\n  <{tag}>TODO</{tag}>"
                                           if varIsNameTODO.get()
                                           else f"  <{tag}>{text}</{tag}>")

            with open(filename, 'w', encoding='UTF8') as fin:
                fin.write("""<?xml version="1.0" encoding="utf-8"?>\n<LanguageData>\n""")
                fin.write('\n'.join(writingTextList))
                if varCollisionOption.get() == 2 and alreadyDefinedDict:  # collision -> merge
                    fin.write("\n\n  <!-- 알파의 추출기는 추출하지 않았지만 이미 존재했던 노드들 \n\n")
                    for tag, text in alreadyDefinedDict.items():
                        if type(text) == list:
                            fin.write(f"  <{tag}>\n")
                            fin.write('\n'.join([f"    <li>{eachText}</li>" for eachText in text]))
                            fin.write(f"\n  </{tag}>\n")
                        else:
                            fin.write(f"  <{tag}>{text}</{tag}>\n")
                    fin.write("  -->")
                fin.write("\n</LanguageData>")

            savedList.append("Keyed")

        # Strings
        for departure in list_strings:
            destination = varDirName.get() + departure.split('Languages\\English')[1]
            if varCollisionOption.get() != 1:  # collision -> not overwrite
                try:
                    with open(destination, 'r', encoding='UTF8') as _:
                        messagebox.showerror("파일 충돌 발견됨",
                                             f"Strings 폴더의 파일이 존재하여 작업을 중단하였습니다.\n" +
                                             f"이미 저장된 파일의 폴더 리스트는 아래와 같습니다.\n{WORD_NEWLINE.join(savedList)}")
                        return
                except FileNotFoundError:
                    pass

            Path('\\'.join(destination.split('\\')[:-1])).mkdir(parents=True, exist_ok=True)
            shutil.copy(departure, destination)

        if list_strings:
            savedList.append("Strings")

        writeConfig(new_exportDir=varDirName.get(), new_exportFile=varFileName.get(),
                    new_isNameTODO=varIsNameTODO.get(), new_collisionOption=varCollisionOption.get())

        messagebox.showinfo("퍼일 저장 완료",
                            "작업이 완료되었습니다.\n새로 작성되거나 변경된 파일의 폴더 리스트는 아래와 같습니다.\n{}".format("\n".join(savedList)))

    Button(frame, text="추출하기", command=exportXml).grid(row=8, column=0)


if __name__ == '__main__':
    goExtractList = []

    dict_class = {}
    dict_keyed = {}
    list_strings = []

    excludes = []
    defaults = []
    includes = []

    excludeHide = []
    defaultHide = []
    includeHide = []

    window = Tk()
    window.title("Alpha의 림월드 모드 언어 추출기")
    window.geometry("800x400+100+100")
    window.iconbitmap(resource_path('icon.ico'))
    Grid.rowconfigure(window, 0, weight=1)
    Grid.columnconfigure(window, 0, weight=1)

    try:
        gameLoc, modsLoc, defExcludes, defIncludes, exDir, exFile, isTODO, colOption = loadConfig()
    except FileNotFoundError:
        writeConfig()
        gameLoc, modsLoc, defExcludes, defIncludes, exDir, exFile, isTODO, colOption = loadConfig()
    except ValueError:
        if messagebox.askyesno("사용자 설정 초기화", "config.dat 파일의 형식이 변경되어 초기화가 필요합니다.\n" +
                                             "초기화 진행 시 사용자가 변경한 설정이 유실됩니다.\n" +
                                             "필요할 경우 초기화를 진행하기 전에 백업해 주세요.\n설정 초기화를 진행할까요?"):
            writeConfig()
        else:
            exit(0)
        gameLoc, modsLoc, defExcludes, defIncludes, exDir, exFile, isTODO, colOption = loadConfig()

    loadInit(window)

    try:
        versionURL = "https://raw.githubusercontent.com/dlgks224/AlphaExtractor/master/CURRENT_VERSION"
        serverVersion = urllib.request.urlopen(versionURL).read().decode("utf-8").replace('\n', '')
        if EXTRACTOR_VERSION != serverVersion:
            if messagebox.askyesno("업데이트 가능",
                                   "새로운 버전의 추출기가 발견되었습니다.\n\n" + \
                                   f"업데이트 버전 : {EXTRACTOR_VERSION} -> {serverVersion}\n\n다운로드 페이지를 열까요?"):
                import webbrowser
                webbrowser.open_new('https://github.com/dlgks224/AlphaExtractor/releases')
                exit(0)
    except (urllib.error.HTTPError, urllib.error.URLError):
        pass

    window.mainloop()
