import os, sys, subprocess, platform, re, requests, xmltodict, openpyxl
from datetime import datetime


deviceAddress = None
userLogin = None
userPwd = None
userLvl = None
adminLogin = None
adminPwd = None
userloginminlen = 2
userloginmaxlen = 8
userpwdminlen = 8
userpwdmaxlen = 16

script_dir = os.path.dirname(os.path.realpath(sys.argv[0]))
xlslistfname = 'addr.xlsx'
xlslistpath = os.path.join(script_dir, xlslistfname)
addrlist = []
allowedUsrLvls = ['Administrator', 'Operator', 'Viewer']
ch_yes = ['yes', 'y']
ch_no = ['no', 'n']
defUserLvl = 'Operator'
useLastKnownData = True

xmlUserTemplate = """<UserList>
<User>
<id>?id?</id>
<userName>?username?</userName>
<password>?userpwd?</password>
<userLevel>?userlvl?</userLevel>
</User>
</UserList>"""

retreadfileexp = 4
storecodesmformat = ''
logchecked = False
date = datetime.today()
logmaxsize = 2097152   # bytes
logmaxdepth = 90       # files
logmaxdepthsingle = 5  # days
logfoldername = 'hikvisuserlog'
folderpath = 'hikvisuserlog'
logfileext = 'log'
logfname = 'current' + ('' if logfileext == '' or logfileext == None else '.' + logfileext)
datetimeformatlog = '[%d/%m/%Y %H:%M:%S] '
datetimeformat = '%Y%m%dT%H%M%S'
dateformat = '%Y%m%d'
timeformat = '%H%M%S'
currentdate = date.strftime(dateformat)
logdatelist = list()

linefeed = '\n'

def ping(host_or_ip, packets=1, timeout=1000):
    ''' Calls system "ping" command, returns True if ping succeeds.
    Required parameter: host_or_ip (str, address of host to ping)
    Optional parameters: packets (int, number of retries), timeout (int, ms to wait for response)
    Does not show any output, either as popup window or in command line.
    Python 3.5+, Windows and Linux compatible
    '''
    # The ping command is the same for Windows and Linux, except for the "number of packets" flag.
    if platform.system().lower() == 'windows':
        command = ['ping', '-n', str(packets), '-w', str(timeout), host_or_ip]
        # run parameters: capture output, discard error messages, do not show window
        result = subprocess.run(command, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, creationflags=0x08000000)
        # 0x0800000 is a windows-only Popen flag to specify that a new process will not create a window.
        # On Python 3.7+, you can use a subprocess constant:
        #   result = subprocess.run(command, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
        # On windows 7+, ping returns 0 (ok) when host is not reachable; to be sure host is responding,
        # we search the text "TTL=" on the command output. If it's there, the ping really had a response.
        return result.returncode == 0 and b'TTL=' in result.stdout
    else:
        command = ['ping', '-c', str(packets), '-w', str(timeout), host_or_ip]
        # run parameters: discard output and error messages
        result = subprocess.run(command, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return result.returncode == 0


# internal: returns text file contents as a list
def readLinesFromFile(filename):
    try:
        file = open(filename, 'r')
        data = file.readlines()
        file.close()
        return data
    except OSError as err:
        logAddLine("Error reading file: " + str(err), ignoreRotation=True)
        return retreadfileexp


# adds val to given list in case its not already stored there
def addUniqueValToList(datalist, val):
    f = True
    for item in datalist:
        if item == val:
            f = False
            break
    if f:
        datalist.append(val)


# stores previous log file to log folder with a timestamp in filename
# deletes logs older than logmaxdepth days
def swapLog(logfolder):
    loglist = next(os.walk(logfolder))[2]
    if len(loglist) >= logmaxdepth:
        try:
            os.remove(folderpath + min(loglist))
        except OSError as err:
            logAddLine("Exception on log file delete: " + str(err), ignoreRotation=True)
    try:
        os.rename(logfname,
                os.path.dirname(logfname) + os.path.sep + currentdate + ('' if logfileext == '' or logfileext == None else '.' + logfileext))
    except OSError as err:
        logAddLine("Exception on log file swap: " + str(err), ignoreRotation=True)


# checks if a current log file doesnt exceed logmaxsize size and doesnt contain
# records for more than logmaxdepthsingle days
def checklog():
    global logfname
    global logchecked
    logfolder = os.path.dirname(logfname)
    if os.path.exists(logfname):
        if os.path.getsize(logfname) < logmaxsize:
            log = readLinesFromFile(logfname)
            if log != retreadfileexp:
                if len(log):
                    for line in log:
                        try:
                            convdate = datetime.strptime(line[:21], datetimeformatlog)
                            addUniqueValToList(logdatelist, convdate.strftime(dateformat))
                        except ValueError:
                            continue
                    c = 0
                    for logdate in logdatelist:
                        for logdatesec in logdatelist:
                            if logdatesec == logdate:
                                c += 1
                    if c >= logmaxdepthsingle:
                        swapLog(os.path.dirname(logfolder + "/"))
        else:
            swapLog(os.path.dirname(logfolder + "/"))
    logchecked = True


# adds line to log file
def logAddLine(logline, ignoreRotation=False, doPrint=True):
    global logchecked
    if not logchecked and not ignoreRotation:
        checklog()
    timestamp = datetime.today().strftime(datetimeformatlog)
    try:
        global logfname
        global storecodesmformat
        file = open(logfname, 'a')
        infostr = timestamp + (storecodesmformat + ' - ' if storecodesmformat != '' else '') + logline + linefeed
        if doPrint:
            print(infostr)
        file.write(infostr)
        file.close()
    except OSError:
        return


class bcolors:
    FORE_RED = '\033[91m'
    FORE_GREEN = '\033[92m'
    FORE_YELLOW = '\033[93m'
    FORE_BLUE = '\033[94m'
    FORE_PURPLE = '\033[95m'
    FORE_CYAN = '\033[96m'
    BOLD = '\033[1m'
    ITALIC = '\033[3m'
    UNDERLINE = '\033[4m'
    RESET = '\x1b[0m'


def fixlink(link):
    if not link.startswith('http://'):
        link = 'http://' + link
    return link


def strhasdigits(teststr):
    return re.search('[0-9]', teststr)


def strhascapitals(teststr):
    return re.search('[A-Z]', teststr)


def strhaschars(teststr):
    return re.search('\w', teststr)


def strnospecial(teststr):
    return not re.search('\W', teststr)


def validateLogin(tgt):
    if userloginmaxlen >= len(tgt) >= userloginminlen and strnospecial(tgt) and strhaschars(tgt):
        return True
    return False


def validatePwd(tgt):
    if userpwdmaxlen >= len(tgt) >= userpwdminlen and strhasdigits(tgt) and strhascapitals(tgt):
        return True
    return False


def validateLvl(tgt):
    tgtf = tgt.replace(tgt.lower()[0], tgt.lower()[0].upper())
    if tgtf in allowedUsrLvls:
        return tgtf
    return None


def loginInput(tgt):
    while True:
        login = input()
        if validateLogin(login):
            return login
        elif login.lower() == 'c':
            sys.exit()
        else:
            print(tgt + ' login must be ' + str(userloginminlen) + '-' + str(userloginmaxlen) +
                  ' latin chars and may contain digits.\nPlease repeat input or enter \'c\' to exit ')


def pwdInput(tgt):
    while True:
        pwd = input()
        if validatePwd(pwd):
            return pwd
        elif pwd.lower() == 'c':
            sys.exit()
        else:
            print(tgt + ' password must be ' + str(userpwdminlen) + '-' + str(userpwdmaxlen) +
                  'latin chars and digits.\nPlease repeat input or enter \'c\' to exit ')


def lvlInput():
    uLvl = input()
    if re.match('[1-3]', uLvl) and len(uLvl) == 1:
        return allowedUsrLvls[int(uLvl) - 1]
    else:
        ulvlf = validateLvl(uLvl)
        if ulvlf is None:
            print('User level must be \'Administrator\', \'Operator\' or \'Viewer\'.\nContinue with \'' + defUserLvl + '\' user level? y/n')
            choice = input().lower()
            if choice in ch_yes:
                return defUserLvl
            else:
                sys.exit()
        else:
            return ulvlf


def datafromxls():
    global deviceAddress, userLogin, userPwd, userLvl, adminLogin, adminPwd, useLastKnownData
    result = True
    if os.path.exists(xlslistpath):
        print(bcolors.FORE_YELLOW + 'Address list found.\nUse last known admin login/pwd in list in case of missing data (y/n/c to exit) ' + bcolors.RESET)
        res = input().lower()
        if res in ch_yes:
            useLastKnownData = True
        elif res in ch_no:
            useLastKnownData = False
        elif res == 'c':
            sys.exit()
        deviceAddress = []
        adminLogin = []
        adminPwd = []
        lkadminlogin = None
        lkadminpwd = None
        wb = openpyxl.load_workbook(xlslistpath)
        sheet = wb.active
        for col in sheet['A']:
            if col.value is None:
                print(bcolors.FORE_RED + 'Error in address list in data file' + bcolors.RESET)
                result = False
                break
            else:
                deviceAddress.append(col.value)
        if result:
            c = 0
            for col in sheet['B']:
                c += 1
                if col.value is None:
                    if useLastKnownData:
                        adminLogin.append(lkadminlogin)
                    else:
                        print(bcolors.FORE_RED + 'Error in admin login list in data file at line ' + str(c) + bcolors.RESET)
                        result = False
                        break
                else:
                    adminLogin.append(col.value)
                    lkadminlogin = col.value
        if result:
            c = 0
            for col in sheet['C']:
                c += 1
                if col.value is None:
                    if useLastKnownData:
                        adminPwd.append(lkadminpwd)
                    else:
                        print(bcolors.FORE_RED + 'Error in admin password list in data file at line ' + str(c) + bcolors.RESET)
                        result = False
                        break
                else:
                    adminPwd.append(col.value)
                    lkadminpwd = col.value
        if result:
            if len(adminLogin) < len(deviceAddress):
                if useLastKnownData:
                    for i in range(0, len(deviceAddress) - len(adminLogin)):
                        adminLogin.append(lkadminlogin)
                else:
                    print(bcolors.FORE_RED + 'Device address and admin login lists length do not match' + bcolors.RESET)
                    result = False
        if result:
            if len(adminPwd) < len(deviceAddress):
                if useLastKnownData:
                    for i in range(0, len(deviceAddress) - len(adminPwd)):
                        adminPwd.append(lkadminpwd)
                else:
                    print(bcolors.FORE_RED + 'Device address and admin password lists length do not match' + bcolors.RESET)
                    result = False
        if result:
            userLogin = sheet['D'][0].value
            userPwd = sheet['E'][0].value
            userLvl = sheet['F'][0].value
    else:
        return False
    return result


def createUsersMultiple():
    global adminLogin, adminPwd, userLogin, userPwd, userLvl
    if not validateLogin(userLogin):
        print('Failed to validate new user login.\nEnter new user login: ')
        logAddLine('Failed to validate new user login', doPrint=False)
        userLogin = loginInput('User')
    if not validatePwd(userPwd):
        print('Failed to validate new user password.\nEnter new user password: ')
        logAddLine('Failed to validate new user password', doPrint=False)
        userPwd = pwdInput('User')
    if not validateLvl(userLvl):
        print('Failed to validate new user level.\nEnter new user level: ')
        logAddLine('Failed to validate new user level', doPrint=False)
        userLvl = lvlInput()
    c = 0
    logAddLine('New user login - ' + userLogin + '; new user pwd - ' + userPwd + '; new user level - ' + userLvl)
    for rec in deviceAddress:
        logAddLine('Starting with device ' + re.sub('http[s]?://', '', rec))
        if not validateLogin(adminLogin[c]):
            print('Failed to validate admin login for ' + rec + '\nEnter device admin login: ')
            logAddLine('Failed to validate admin login for ' + rec, doPrint=False)
            adminLogin = loginInput('Device administrator')
        if not validatePwd(adminPwd[c]):
            print('Failed to validate admin password for ' + rec + '\nEnter device admin password: ')
            logAddLine('Failed to validate admin password for ' + rec, doPrint=False)
            adminPwd = pwdInput('Device administrator')
        if not ping(re.sub('http[s]?://', '', rec)):
            logAddLine(rec + ' - unreachable by ICMP ping')
            continue
        addPort = False
        addrstr = fixlink(rec) + ':8080'
        sess = requests.Session()
        sess.auth = (adminLogin[c], adminPwd[c])
        try:
            sess.post(addrstr)
        except requests.exceptions.ConnectionError:  # requests.exceptions.InvalidSchema:
            addPort = True
        if addPort:
            addrstr = fixlink(rec) + ':80'
            try:
                sess.post(addrstr)
            except requests.exceptions.ConnectionError:
                logAddLine('Failed to connect to ' + rec + ' on ports 80 and 8080')
                continue
        resp = sess.get(addrstr + '/ISAPI/Security/users')
        userdict = xmltodict.parse(resp.content.decode('utf-8'))
        if 'html' in userdict and 'Unauthorized' in userdict['html']['head']['title']:
            logAddLine('Could not authorize at ' + rec + ' with provided credentials')
            continue
        exuquan = 0
        skipCreate = False
        for usr in userdict['UserList']['User']:
            if usr['userName'] == userLogin:
                logAddLine('Failed to create user at ' + rec + ' - user already exists')
                c += 1
                skipCreate = True
                break
            if 'id' in usr:
                exuquan += 1
        if skipCreate:
            continue
        exuquan += 1
        usrxml = xmlUserTemplate
        usrxml = usrxml.replace('?id?', str(exuquan))
        usrxml = usrxml.replace('?username?', userLogin)
        usrxml = usrxml.replace('?userpwd?', userPwd)
        usrxml = usrxml.replace('?userlvl?', userLvl)
        resp = sess.post(addrstr + '/ISAPI/Security/users', data=usrxml)
        devans = xmltodict.parse(resp.content.decode('utf-8'))['ResponseStatus']['statusString']
        if devans == 'OK':
            logAddLine('User ' + userLogin + ' created at ' + rec)
        else:
            logAddLine('Could not create user at ' + rec + ' - ' + devans)
        c += 1


def createUserSingle():
    global deviceAddress, adminLogin, adminPwd, userLogin, userPwd, userLvl
    try:
        deviceAddress = fixlink(sys.argv[1])
        userLogin = sys.argv[2]
        userPwd = sys.argv[3]
        userLvl = sys.argv[4]
        adminLogin = sys.argv[5]
        adminPwd = sys.argv[6]
    except IndexError:
        pass
    if deviceAddress is None:
        print('Enter device address: ')
        deviceAddress = fixlink(input().lower())
    if userLogin is None:
        print('Enter user login: ')
        userLogin = loginInput('User')
    if userPwd is None:
        print('Enter user password: ')
        userPwd = pwdInput('User')
    if userLvl is None:
        print('Enter user level (\'Administrator\'[1], \'Operator\'[2] or \'Viewer\'[3]): ')
        userLvl = lvlInput()
    if adminLogin is None:
        print('Enter device administrator login: ')
        adminLogin = loginInput('Device administrator')
    if adminPwd is None:
        print('Enter device administrator password: ')
        adminPwd = pwdInput('Device administrator')
    logAddLine('New user login - ' + userLogin + '; new user pwd - ' + userPwd + '; new user level - ' + userLvl)
    logAddLine('Starting with device ' + re.sub('http[s]?://', '', deviceAddress))
    if not ping(re.sub('http[s]?://', '', deviceAddress)):
        logAddLine(deviceAddress + ' - unreachable by ICMP ping')
        return
    exuquan = 0
    addPort = False
    addrstr = deviceAddress + ':8080'
    sess = requests.Session()
    sess.auth = (adminLogin, adminPwd)
    try:
        sess.post(addrstr)
    except IndexError:
        pass
    except requests.exceptions.ConnectionError:  # requests.exceptions.InvalidSchema:
        addPort = True
    if addPort:
        addrstr = deviceAddress + ':80'
        try:
            sess.post(addrstr)
        except requests.exceptions.ConnectionError:
            logAddLine('Failed to connect to ' + deviceAddress + 'on ports 80 and 8080')
            return
    resp = sess.get(deviceAddress + '/ISAPI/Security/users')
    userdict = xmltodict.parse(resp.content.decode('utf-8'))
    if 'html' in userdict and 'Unauthorized' in userdict['html']['head']['title']:
        logAddLine('Could not authorize at ' + deviceAddress + ' with provided credentials')
        return
    for rec in userdict['UserList']['User']:
        if rec['userName'] == userLogin:
            logAddLine('Failed to create user at ' + deviceAddress + ' - user already exists')
            sys.exit()
        if 'id' in rec:
            exuquan += 1
    exuquan += 1
    usrxml = xmlUserTemplate
    usrxml = usrxml.replace('?id?', str(exuquan))
    usrxml = usrxml.replace('?username?', userLogin)
    usrxml = usrxml.replace('?userpwd?', userPwd)
    usrxml = usrxml.replace('?userlvl?', userLvl)
    resp = sess.post(deviceAddress + '/ISAPI/Security/users', data=usrxml)
    devans = xmltodict.parse(resp.content.decode('utf-8'))['ResponseStatus']['statusString']
    if devans == 'OK':
        logAddLine('User ' + userLogin + ' created at ' + deviceAddress)
    else:
        logAddLine('Could not create user at ' + deviceAddress + ' - ' + devans)


# SCRIPT
folderpath = os.path.dirname(os.path.realpath(__file__)) + os.path.sep
logfolderpath = folderpath + logfoldername + os.path.sep
if not os.path.exists(logfolderpath):
    os.makedirs(logfolderpath)
logfname = logfolderpath + logfname
logAddLine('Script started')
if datafromxls():
    logAddLine('Data file found; creating multiple users', doPrint=False)
    createUsersMultiple()
else:
    logAddLine('Creating a single user', doPrint=False)
    createUserSingle()
logAddLine('Script finished')

