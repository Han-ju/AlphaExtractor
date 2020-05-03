import glob
import os
import shutil
import xml.etree.ElementTree as et
from pathlib import Path
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

EXTRACTABLE_DIRS = ["Defs", "Languages"]
CONFIG_VERSION = 2
EXTRACTOR_VERSION = "0.8.3"
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


def loadConfig(fileName='config.dat'):
    with open(fileName, 'r', encoding='UTF8') as fin:
        configs = fin.read().split('\n')

    if CONFIG_VERSION != int(configs[3]):
        raise ValueError()

    modDir = configs[5]
    definedExcludes = configs[7].replace(' ', '').split('/')
    definedIncludes = configs[9].replace(' ', '').split('/')
    exportDir = configs[11]
    exportFile = configs[13]
    isNameTODO = (configs[15] == 'True')
    collisionOption = int(configs[17])

    return modDir, definedExcludes, definedIncludes, exportDir, exportFile, isNameTODO, collisionOption


# noinspection PyShadowingNames
def writeConfig(new_modDir=None, new_exportDir=None, new_exportFile=None,
                new_isNameTODO=None, new_collisionOption=None, fileName="config.dat"):
    try:
        modDir, definedExcludes, definedIncludes, exportDir, exportFile, isNameTODO, collisionOption = loadConfig()
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
    except ValueError:
        modDir = ""
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

    for i in range(3):
        Grid.rowconfigure(frame, i, weight=1)
    for i in range(2):
        Grid.columnconfigure(frame, i, weight=1)

    labelText = f"""림월드 모드들이 위치한 폴더를 선택해 주십시오

해당 폴더에는 모드 폴더들이 존재해야 합니다.
스팀 창작마당 모드의 경우 기본값은 다음과 같습니다.

[C:/Program Files (x86)/Steam/steamapps/workshop/content/294100]

이 텍스트를 클릭할 경우 위 경로가 적용됩니다.

현재 추출 가능한 폴더: Defs, Keyed, Strings
모든 종류의 개선안이나, 버그 제보는 항상 환영합니다.

version = {EXTRACTOR_VERSION}"""

    label = Label(frame, text=labelText)
    label.grid(row=0, column=0, columnspan=4, sticky=N + S + E + W)

    # noinspection PyUnusedLocal
    def onClickLabel(evt):
        entry.delete(0, END)
        entry.insert(0, "C:/Program Files (x86)/Steam/steamapps/workshop/content/294100")

    label.bind("<Button-1>", onClickLabel)

    entry = Entry(frame)
    entry.insert(0, modsLoc)
    entry.grid(row=1, column=0, columnspan=3, sticky=E + W)

    def onClick1():
        entry.delete(0, END)
        entry.insert(0, filedialog.askdirectory())

    btn = Button(frame, text="폴더 선택", command=onClick1)
    btn.grid(row=1, column=3)

    def onClick2():
        modsLocation = entry.get()
        if not modsLocation:
            messagebox.showerror("경로 선택되지 않음", "모드들이 위치한 폴더를 선택해 주십시오.")
            return
        writeConfig(new_modDir=modsLocation)
        frame.destroy()
        loadSelectMod(window)

    btn = Button(frame, text="확인", command=onClick2)
    btn.grid(row=2, column=0, columnspan=4)


def loadSelectMod(window):
    frame = Frame(window)
    frame.grid(row=0, column=0, sticky=N + S + E + W)
    Grid.rowconfigure(frame, 1, weight=1)
    Grid.columnconfigure(frame, 0, weight=2)
    Grid.columnconfigure(frame, 1, weight=1)

    textLabel1 = Label(frame, text="추출할 모드를 선택하세요")
    textLabel1.grid(row=0, column=0, sticky=N + S + E + W, padx=5, pady=5)

    modListBoxValue = StringVar()
    modListBox = Listbox(frame, selectborderwidth=3, listvariable=modListBoxValue)
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

    modPathList = [modPath.replace('\\', '/') for modPath in glob.glob(modsLoc + '/*') if os.path.isdir(modPath)]
    if not modPathList:  # If No Mod
        messagebox.showerror("모드 폴더 찾을 수 없음", "선택한 폴더에 어떤 하위 폴더도 존재하지 않습니다.\n프로그램을 종료합니다.")
        exit(0)

    modsNameDict = {}
    sep = ' | '
    for modPath in modPathList:
        try:
            name = et.parse(modPath + '/About/About.xml').getroot().find('name').text
            modsNameDict[f"{int(modPath.split('/')[-1]):010d}{sep}{name}"] = modPath
        except (FileNotFoundError, ValueError):
            modsNameDict[f"0000000000{sep}{modPath.split('/')[-1]}"] = modPath
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


def parse_recursive(parent, className, tag, lastTag=None):
    if list(parent):
        tags = [child.tag == 'li' for child in list(parent)]
        texts = [child.text for child in list(parent)]
        if False and all(tags) and all(texts) and tags and texts: ##########################################################
            yield className, lastTag, tag, texts
        else:
            num_list = 0
            for child in list(parent):
                if child.tag == 'li':
                    yield from parse_recursive(child, className, tag + '.' + str(num_list), lastTag)
                    num_list += 1
                else:
                    yield from parse_recursive(child, className, tag + '.' + child.tag, child.tag)
    else:
        yield className, lastTag, tag, (parent.text if parent.text else "ERROR:{$BLANK TEXT}")


def extractDefs(root):
    if 'value' != root.tag and 'Defs' != root.tag:
        raise ValueError("첫 태그가 Defs가 아닙니다. 오류를 발생시킨 모드 이름과 파일명을 Alpha에게 제보해주세요.")

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
                    dict_keyed[node.tag] = node.text
            list_strings = glob.glob(goExtract + "\\English\\Strings\\**\\*.txt", recursive=True)
        else:
            messagebox.showerror("에러 발생", "Patches 등 Defs, Keyed 이외의 모든 폴더는 아직 추출할 수 없습니다.\n자동으로 제외합니다.")

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
    easterEgg = 0

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
    Grid.rowconfigure(window, 0, weight=1)
    Grid.columnconfigure(window, 0, weight=1)

    try:
        modsLoc, defExcludes, defIncludes, exDir, exFile, isTODO, colOption = loadConfig()
    except FileNotFoundError:
        writeConfig()
        modsLoc, defExcludes, defIncludes, exDir, exFile, isTODO, colOption = loadConfig()
    except ValueError:
        if messagebox.askyesno("사용자 설정 초기화", "config.dat 파일의 형식이 변경되어 초기화가 필요합니다.\n" +
                                             "초기화 진행 시 사용자가 변경한 설정이 유실됩니다.\n" +
                                             "필요할 경우 초기화를 진행하기 전에 백업해 주세요.\n설정 초기화를 진행할까요?"):
            writeConfig()
        else:
            exit(0)
        modsLoc, defExcludes, defIncludes, exDir, exFile, isTODO, colOption = loadConfig()

    loadInit(window)

    window.mainloop()
