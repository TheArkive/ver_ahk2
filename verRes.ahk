; ===============================================================
; Example
; ===============================================================

; msg := "*** WARNING ***`n"
     ; . "When you try this, make a COPY of AutoHotkey.exe!!!`n"
     ; . "I suggest you don't try this on the original.`n`n"
     ; . "Do you want to continue?"

; If Msgbox(msg,,4) = "No"
    ; ExitApp

; sFile := FileSelect("1",A_ScriptDir)
; If !sFile
    ; ExitApp

; vi := ver(sFile) ; specify only a file to get an automatic list of Version resources

; list := ""
; For i, obj in vi.ResList ; version resource names are usually numbers
    ; list .= (list?"`n":"") obj[1] " / " obj[2]

; msgbox "Version Resource List (Name / Lang):`n`n" list

; vi := ver(sFile,1) ; Now recreating the obj, and specifying Resource #1 on creation.  Usually the version resource name is 1.

;;;;;;; To use this, uncomment ReadBuffer() func at the bottom
;;;;;;; msgbox ReadBuffer(vi,vi.OrigObj.ptr) ; for debugging ; check original object

; arr := vi.ListValues()
; list := ""
; For i, obj in arr
    ; list .= (list?"`n":"") obj.key " = " obj.Value
; msgbox "List all StringTable properties:`n`n" list

; msgbox "Reading file version data:`n`n"
     ; . "File Desc:`t" vi.FileDescription "`n" ; read data as properties
     ; . "File Ver:`t`t" vi.FileVersion

; vi.FileDescription("New File Description")
; vi.FileVersion("1.2.3.4") ; write data as methods
; vi.ProductName("My Product Name")
; vi.ProductVersion("4.3.2.1")
; vi.Comments("test comment")

; vi.Apply() ; apply changes and construct buffer
;;;;;;; To use this, uncomment ReadBuffer() func at the bottom
;;;;;;; msgbox ReadBuffer(vi,vi.outObj.ptr) ; for debugging ; check output buffer before update

; vi.BeginUpdate()
; vi.Save()
; result := vi.EndUpdate()
; msgbox "Success:`t" (result?"true":"false") " / Last Error: " A_LastError "`n`n"
     ; . "Check the file properties now!"

; ===============================================================
; NOTE:  I plan to make a suite of these resource classes meant
; to be used with each other for performing resource udpates.
; This is the first one.  Icons and RCDATA are next.
;
; This class is not designed to handle multiple StringTables
; or multiple Translation members.  The reson for this is because
; the MS docs site actually suggest having one resource per
; language, so this class conforms to this recommendation.
;
; I may update this class to handle more complex version resources
; in the future, but this will not be a priority.
; ===============================================================
; ver class
;
;   Create object:
;       vi := ver(file_name [ , resource_name := ""] )
;
;       If you specify a resource name, that resource is automatically loaded and
;       parsed.  If not, then you get a list of Version resources.
;
;   Get property data:
;       vi.prop_name
;
;       Ex:
;       vi.FileDescription
;
;   Write property data:
;       vi.prop_name(new_value)
;
;       Ex:
;       vi.FileDescription("new file description")
;
;   Properties:
;
;       vi.ResList  = Array of Version resource names.  Each element in the array
;                     is an array with 2 members:
;                       arr[1] = resource name
;                       arr[2] = resource language (ie. 1033 for English)
;
;       vi.cp       = The codepage value.  You can change this if needed.
;
;       vi.lang     = The language value.  You can change this if needed.
;           * See codepage and lang values below.
;           * When you use vi.Apply() method, these values will be referenced to
;             construct the StringTable szKey and the Translation member value.
;
;       vi.hUpdate  = The handle to the file for the update operation.
;                     This is mostly useful if doing other resource updates.
;                     You will need to pass it to other classes/functions to
;                     do other resource updates.
;
;   * These properties below may be useful for advanced purposes:
;
;       vi.origObj  = The original version resource
;       vi.outObj   = The obj used for output/writing new version resource.
;       vi.oldLang  = This is used internally to remove the initial resource before
;                     writing the new resource.  Referencing this property is not 
;                     usually very helpful.
;
;       vi.children / vi.StrTable = These are arrays of elements in the version
;                                   resource data.  You can parse these and edit
;                                   or inspect them manually if you wish.
;
;                                   * Useful element properties:
;
;                                   szKey, value
;
;                                   Ex:
;
;                                   vi.StrTable[3].value    = gets the value of the 3rd item
;                                   vi.StrTable[2].szKey    = gets the szKey (field title) of the 2nd item
;
;       vi.outList  = This is the output array used to construct the output object.
;                     It is structured similarly to the children and StrTable properties,
;                     except it is a single array, and all elements are in the proper order.
;   Methods:
;
;       Load(resource_name := 1)
;       vi.Load(1)  = Loads the specified Version resource name.  If no name is specified
;                     then it defaults to Version resource 1.
;
;       List()
;       vi.List()   = Populates the vi.ResList property with an array of Version resource
;                     names.  If only a file is specified on creation of vi object, then
;                     this list is automatically populated.
;
;       Apply()
;       vi.Apply()  = Constructs the new version resource and saves it into vi.outObj
;
;       BeginUpdate()
;       hUpdate := vi.BeginUpdate() = Preps the file for update. You don't need
;                                     to do this if it has already been done,
;                                     ie, if you are updating multiple resources.
;
;                                     The output is the handle to the file for
;                                     the update operation.  This handle is also
;                                     stored in vi.hUpdate
;
;       Save()
;       bool := vi.Save(hUpdate:=0) = Performs the resource update.  Returns bool,
;                                     True = success / False = failure
;                                     You can specify an external hUpdate handle.
;                                     Otherwise the internally recorded handle is
;                                     used if it exists.
;
;       EndUpdate()
;       bool := vi.EndUpdate(hUpdate:=0) = Commits the update changes to the file.
;                                          The return value indicates success.
;                                          True = success / False = failure
;                                          You can specify an external hUpdate handle.
;                                          Otherwise the internally recorded handle is
;                                          used if it exists.
;
;                                          After you use this method, you should normally
;                                          destroy / ignore / overwrite the class instance.
;
;
;       ListValues()
;       arr := vi.ListValues() = This loads all szKey's and the associated string table value.
;                                The output is an array.  Each element has the following
;                                properties:
;
;                                   obj.key
;                                   obj.value
; ===============================================================
class ver { ; thanks to VersionRes.ahk for inspiration (lib from Ahk2Exe)
    sFile := "", name := "", hUpdate := 0
    children := [], strTable := [], outList := []
    origObj := "", outObj := ""
    EnumCb := {}
    names := [], ResList := []
    lang := "", cp := "", oldLang := ""
    __New(sFile, name:="") {
        this.sFile := sFile, this.name := name
        
        If IsInteger(name) && (name) {
            this.Load(name)
            this.Read(this.origObj.ptr)
        } Else
            this.List()
    }
    EnumRes(hModule, sType, p*) { ; sName, [Lang,] lParam
        name := (((sName:=p[1])>>16)=0) ? sName : StrGet(sName)
        (p.Length = 2) ? this.names.Push(name) : this.ResList.Push( [name,p[2]] ) ; version resource [name, lang]
        return true
    }
    Load(name:=1) {
        name := IsInteger(name) ? "#" name : name
        hModule := DllCall("LoadLibraryEx","Str",(this.sFile), "UPtr", 0, "UInt", 0x2, "UPtr")
        fRsc := DllCall("FindResource","UPtr",hModule,"Str",name,"Ptr",16,"UPtr")
        hRsc := DllCall("LoadResource","UPtr",hModule,"UPtr",fRsc,"UPtr")   ; resource handle
        ptr := DllCall("LockResource","UPtr",hRsc,"UPtr")                   ; resource pointer
        
        this.origObj := Buffer(size := DllCall("SizeofResource","UPtr",hModule,"UPtr",fRsc,"UInt"))
        DllCall("RtlCopyMemory", "UPtr", this.origObj.ptr, "UPtr", ptr, "UPtr", size)
        r1 := DllCall("FreeLibrary","UPtr",hModule,"UPtr")
    }
    List() {
        If !(hModule := DllCall("LoadLibrary","Str",this.sFile, "UPtr"))
            return false
        
        this.EnumCb.fnc := ObjBindMethod(this,"EnumRes")
        
        Loop (p:=2) {
            this.EnumCb.ptr := CallbackCreate(this.EnumCb.fnc,"F",p+2)
            
            If (A_Index=1) {
                r1 := DllCall("EnumResourceNames", "UPtr", hModule, "Ptr", 16, "UPtr", this.EnumCb.ptr, "UPtr", 0, "Int")
            } Else {
                For i, name in this.names
                    DllCall("EnumResourceLanguagesEx", "UPtr", hModule, "Ptr", 16, "Ptr", name, "UPtr", this.EnumCb.ptr
                                                     , "UPtr", 0, "UInt", 0x1, "UShort", 0)
            }
            p++, CallbackFree(this.EnumCb.ptr)
        }
        r1 := DllCall("FreeLibrary","UPtr",hModule)
    }
    Read(curAddr) { ; used internally
        this.children := [], this.strTable := [], this.outList := [] ; reset lists
        this.origObj := "", this.outObj := ""
        limit := curAddr + NumGet(curAddr,"UShort")
        
        While (curAddr < limit) {
            obj := {wLength:NumGet(curAddr,"UShort")
                   ,wValLen:(valLen := NumGet(curAddr,2,"UShort"))
                   ,wType:(wType := NumGet(curAddr,4,"UShort"))
                   ,szKey:(szKey := StrGet(curAddr+6))
                   ,value:(valLen&&!wType)?Buffer(valLen):""} ; create buffer for NON-text only
            
            curAddr := (curAddr + 6 + StrPut(szKey) + 3) & ~3
            If (valLen && !wType)
                DllCall("RtlCopyMemory", "UPtr", obj.value.ptr, "UPtr", curAddr, "UPtr", valLen)
            Else If (valLen && wType)
                obj.value := StrGet(curAddr)
            
            If (A_Index=3)
                (this.lang := this.oldLang := SubStr(szKey,1,4)), this.cp := SubStr(szKey,5)
            
            (A_Index>=4) ? this.strTable.Push(obj) : this.children.Push(obj)
            curAddr := (curAddr + ((wType&&valLen)?StrPut(obj.value):valLen) + 3) & ~3
        }
        
        this.children.Push(this.strTable.RemoveAt(this.strTable.Length)) ; move last two elements to this.children
        this.children.InsertAt(this.children.Length, this.strTable.RemoveAt(this.strTable.Length))
    }
    ListValues() {
        list := []
        For i, obj in this.StrTable
            list.Push({key:obj.szKey, value:obj.value})
        return list
    }
    __Get(prop, p) {
        found := false, o := ""
        For i, obj in this.strTable
            If (found := (prop = obj.szKey) && (o := obj))
                Break
        return found ? o.value : ""
    }
    __Call(prop, p) {
        found := false, o := "", value := p[1]
        For i, obj in this.strTable
            If (found := (prop = obj.szKey) && (item := i))
                Break
        If found {
            o := this.strTable[item]
            o.value := value
            o.wLength := ((6 + StrPut(prop) + 3) & ~3) + StrPut(value)
            o.wValLen := Round(StrPut(value)/2)
        } Else {
            this.children.Push({wLength:((6 + StrPut(prop) + 3) & ~3) + StrPut(value)
                               ,wValLen:(StrPut(value)/2),wType:1,szKey:prop,value:value})
        }
    }
    Apply() { ; Read elements and recalc sizes, and then construct output buffer (this.outObj)
        (strTableSize := 0), (o := this.children), (o[3].szKey := this.lang . this.cp)
        For i, obj in this.strTable
            strTableSize += (obj.wLength + 3) & ~3
        
        o[3].wLength := strTableSize += (6 + 18)    ; Update StringTable size
        o[2].wLength := o[3].wLength + 6 + 30       ; Update StringFileInfo size
        o[1].wLength := ((6 + 32 + 3) & ~3) + 52 + o[2].wLength + 68 ; Update VS_VERSION_INFO
        
        NumPut("UInt", ((("0x" this.cp)<<16) | ("0x" this.lang)), o[5].value) ; Write Translation value
        
        Loop 3 ; re-order elements into one array (this.outList)
            this.outList.Push(o[A_Index])
        For i, obj in this.strTable
            this.outList.Push(obj)
        Loop 2
            this.outList.Push(o[A_Index+3])
        
        fvArr := StrSplit(this.FileVersion,"."), pvArr := StrSplit(this.ProductVersion,".") ; update VS_FIXEDFILEINFO
        NumPut("UInt", (fvArr.Has(1)?fvArr[1]<<16:0) | (fvArr.Has(2)?fvArr[2]:0), (obj := o[1].value), 8)
        NumPut("UInt", (fvArr.Has(3)?fvArr[3]<<16:0) | (fvArr.Has(4)?fvArr[4]:0), obj, 12)
        NumPut("UInt", (pvArr.Has(1)?pvArr[1]<<16:0) | (pvArr.Has(2)?pvArr[2]:0), obj, 16)
        NumPut("UInt", (pvArr.Has(3)?pvArr[3]<<16:0) | (pvArr.Has(4)?pvArr[4]:0), obj, 20)
        
        curAddr := (buf := Buffer(this.outList[1].wLength)).ptr
        
        For i, obj in this.outList {
            NumPut("UShort",obj.wLength,curAddr)
            NumPut("UShort",valLen := obj.wValLen,curAddr,2)
            NumPut("UShort",wType := obj.wType,curAddr,4)
            StrPut(obj.szKey, curAddr+6)
            
            curAddr := (curAddr + 6 + StrPut(obj.szKey) + 3) & ~3
            
            If (obj.wValLen && !obj.wType) ; copy buffer data (VS_FIXEDFILEINFO or Translation only)
                DllCall("RtlCopyMemory", "UPtr", curAddr, "UPtr", obj.value.ptr, "UPtr", obj.value.size)
            Else If (obj.wValLen && obj.wType)
                StrPut(obj.value, curAddr) ; normally for String structure values
            
            curAddr := (curAddr + ((wType&&valLen)?StrPut(obj.value):valLen) + 3) & ~3
        }
        
        this.outObj := buf ; save output buffer for resource update
    }
    BeginUpdate() {
        return (this.hUpdate := DllCall("BeginUpdateResource", "Str", this.sFile, "UInt", 0, "UPtr"))
    }
    Save(hUpdate:=0) {
        If !(hUpdate := (hUpdate)?hUpdate:this.hUpdate)
            return false
        
        If !DllCall("UpdateResource", "Ptr", hUpdate, "Ptr", 16, "Ptr", this.name ; delete old lang resource
                                    , "UShort", ("0x" this.oldLang), "UPtr", 0, "UInt", 0)
            return false
            
        return DllCall("UpdateResource", "Ptr", hUpdate, "Ptr", 16, "Ptr", this.name ; write new resource
                                       , "UShort", ("0x" this.lang), "Ptr", this.outObj.ptr, "UInt", this.outObj.size)
    }
    EndUpdate(hUpdate:=0) {
        If !(hUpdate := (hUpdate)?hUpdate:this.hUpdate)
            return false
        return DllCall("EndUpdateResource","UPtr",hUpdate,"Int",(this.hUpdate := 0))
    }
}

; ===============================================================
; Common Version Info Fields
; ===============================================================
; Comments
; CompanyName
; FileDescription
; FileVersion
; InternalName
; LegalCopyright
; LegalTrademarks
; OriginalFilename
; PrivateBuild
; ProductName
; ProductVersion
; SpecialBuild

; ==================================================================
; VersionInfo LANG IDs
; ==================================================================
; Code    Language                Code    Language
; 0x0401  Arabic                  0x0415  Polish
; 0x0402  Bulgarian               0x0416  Portuguese (Brazil)
; 0x0403  Catalan                 0x0417  Rhaeto-Romanic
; 0x0404  Traditional Chinese     0x0418  Romanian
; 0x0405  Czech                   0x0419  Russian
; 0x0406  Danish                  0x041A  Croato-Serbian (Latin)
; 0x0407  German                  0x041B  Slovak
; 0x0408  Greek                   0x041C  Albanian
; 0x0409  U.S. English (1033)     0x041D  Swedish
; 0x040A  Castilian Spanish       0x041E  Thai
; 0x040B  Finnish                 0x041F  Turkish
; 0x040C  French                  0x0420  Urdu
; 0x040D  Hebrew                  0x0421  Bahasa
; 0x040E  Hungarian               0x0804  Simplified Chinese
; 0x040F  Icelandic               0x0807  Swiss German
; 0x0410  Italian                 0x0809  U.K. English
; 0x0411  Japanese                0x080A  Spanish (Mexico)
; 0x0412  Korean                  0x080C  Belgian French
; 0x0413  Dutch                   0x0C0C  Canadian French
; 0x0414  Norwegian ? Bokmal      0x100C  Swiss French
; 0x0810  Swiss Italian           0x0816  Portuguese (Portugal)
; 0x0813  Belgian Dutch           0x081A  Serbo-Croatian (Cyrillic)
; 0x0814  Norwegian ? Nynorsk

;===============================================
; VersionInfo CHARSETs
;===============================================
; Decimal Hexadecimal Character Set
; 0       0000        7-bit ASCII
; 932     03A4        Japan (Shift ? JIS X-0208)
; 949     03B5        Korea (Shift ? KSC 5601)
; 950     03B6        Taiwan (Big5)
; 1200    04B0        Unicode
; 1250    04E2        Latin-2 (Eastern European)
; 1251    04E3        Cyrillic
; 1252    04E4        Multilingual
; 1253    04E5        Greek
; 1254    04E6        Turkish
; 1255    04E7        Hebrew
; 1256    04E8        Arabic




; ===========================================================================
; Support func to parse in/out version objects and display data only.
; Not needed for normal usage.
; ===========================================================================
; ReadBuffer(vi,curAddr) {
    ; limit := curAddr + (fm := NumGet(curAddr,"UShort"))
    ; msg := "" ; for debugging / visualizing data
    ; VS_FIXEDFILEINFO := "", Translation := ""
    
    ; While curAddr < limit {
        ; obj := {wLength:NumGet(curAddr,"UShort")
               ; ,wValLen:(valLen := NumGet(curAddr,2,"UShort"))
               ; ,wType:(wType := NumGet(curAddr,4,"UShort"))
               ; ,szKey:(szKey := StrGet(curAddr+6))
               ; ,value:(valLen&&!wType)?Buffer(valLen):""} ; create buffer for NON-text only
        
        ; curAddr := (curAddr + 6 + StrPut(szKey) + 3) & ~3
        ; If (valLen && !wType)
            ; DllCall("RtlCopyMemory", "UPtr", obj.value.ptr, "UPtr", curAddr, "UPtr", valLen)
        ; Else If (valLen && wType)
            ; obj.value := StrGet(curAddr)
        
        ; test_val1 := ((Type(obj.value)="Buffer") ? "obj" : obj.value)
        ; test_val2 := ((Type(obj.value)="Buffer") ? obj.value.size : (obj.value?StrPut(obj.value):""))
        ; msg .= "wLength: " obj.wLength "`n"
             ; . "wValLen: " obj.wValLen "`n"
             ; . "wType: " obj.wType "`n"
             ; . "szKey: " obj.szKey " (" StrPut(obj.szKey) ")`n"
             ; . "value: " test_val1 . " (" test_val2 . ")`n"
             ; . "====================================`n`n"
        
        ; If (obj.szKey = "VS_VERSION_INFO") {
            ; Loop 13
                ; VS_FIXEDFILEINFO .= (VS_FIXEDFILEINFO?"`n":"") Format("0x{:08X}",NumGet(obj.value,(A_Index-1)*4,"UInt"))
        ; } Else If (obj.szKey = "Translation")
            ; Translation := Format("0x{:08X}",NumGet(obj.value,"UInt"))
        
        ; curAddr := (curAddr + ((wType&&valLen)?StrPut(obj.value):valLen) + 3) & ~3
    ; }
    
    ; msg .= "VS_FIXEDFILEINFO`n======================`n" VS_FIXEDFILEINFO "`n`n"
         ; . "Translation`n=====================`n" Translation
    ; return msg
; }

; dbg(_in) {
    ; Loop Parse _in, "`n", "`r"
        ; OutputDebug "AHK: " A_LoopField
; }