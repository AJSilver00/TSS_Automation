class email:
    import pyautogui
    import time
    import os
    import pyperclip
    import keyboard

    def __init__(self):
        self.os.chdir(r'C:\Users\Aimee\Pictures\TSS Python Screenshots\proforma reminder')
        self.jobref = None
        self.jobreffound = False
        self.invoicingtabunsel = False
        self.invoicingtabsel = False
        self.invoiceproductionsel = False
        self.dropdown = False
        self.orderconfprofinv = False
        self.aspdfdropdown = False
        self.emailaspdf = False
        self.emailto = False
        self.jobcontactemail = False
        self.emails2 = False
        self.displayasemail = False
        self.popup = False
        self.proformaemailwindow = False
        self.TCs = False

    def r_proforma(self):
        detailstabsel = self.pyautogui.locateOnScreen('detailstab.png', confidence=0.9)
        detailstabsel = self.pyautogui.center(detailstabsel)
        x, y = detailstabsel
        self.pyautogui.click(x, y)
        print('I have selected the Details tab.')
        while self.jobreffound == False:
            jobref = self.pyautogui.locateOnScreen('jobref.png', confidence=0.9)
            if jobref != None:
                self.pyautogui.hotkey('ctrl', 'c')
                self.jobref = self.pyperclip.paste()
                print('Job ref {} reminder will be sent now.'.format(self.jobref))
                self.jobreffound = True
                break
        while self.invoicingtabsel == False:
            invoicingtabunsel = self.pyautogui.locateOnScreen('invoicingtabunsel.png', confidence=0.9)
            invoicingtabsel = self.pyautogui.locateOnScreen('invoicingtabsel.png', confidence=0.9)
            if invoicingtabunsel != None:
                invoicingtabunsel = self.pyautogui.center(invoicingtabunsel)
                x, y = invoicingtabunsel
                self.pyautogui.click(x, y)
                print('I have selected the Invoicing tab')
                self.invoicingtabsel = True
            else:
                invoicingtabsel = self.pyautogui.locateOnScreen('invoicingtabsel.png', confidence=0.9)
                if invoicingtabsel != None:
                    print('Invoicing tab is already selected')
                    self.invoicingtabsel = True
                    break
        while self.invoiceproductionsel == False:
            invoiceproductionunsel = self.pyautogui.locateOnScreen('invoiceproductionunsel.png', confidence=0.9)
            invoiceproductionsel = self.pyautogui.locateOnScreen('invoiceproductionsel.png', confidence=0.9)
            if invoiceproductionunsel != None:
                invoiceproductionunsel = self.pyautogui.center(invoiceproductionunsel)
                x, y = invoiceproductionunsel
                self.pyautogui.click(x, y)
                print('Invoice Production tab is now selected')
                self.invoiceproductionsel == True
                break
            else:
                if invoiceproductionsel != None:
                    print('Invoice Production tab is already selected')
                    self.invoiceproductionsel == True
                    break
                else:
                    print('Invoice Production tab has not loaded yet. Please wait.')
        while self.dropdown == False:
            mainhireinvb4payment = self.pyautogui.locateOnScreen('mainhireinvb4payment.png', confidence=0.9)
            if mainhireinvb4payment != None:
                x = mainhireinvb4payment.left + mainhireinvb4payment.width - 20
                y = mainhireinvb4payment.top + 20
                self.pyautogui.click(x, y)
                self.dropdown = True
                break
        while self.orderconfprofinv == False:
            orderconfprofinv = self.pyautogui.locateOnScreen('orderconfprofinv.png', confidence=0.9)
            if orderconfprofinv != None:
                orderconfprofinv = self.pyautogui.center(orderconfprofinv)
                x, y = orderconfprofinv
                self.pyautogui.click(x, y)
                print('I have selected the Order Confirmation dropdown option')
                self.orderconfprofinv = True
                break
            else:
                print('Dropdown list not yet loaded. Please wait.')
        while self.aspdfdropdown == False:
            aspdfdropdown = self.pyautogui.locateOnScreen('aspdfdropdown.png', confidence=0.9)
            if aspdfdropdown != None:
                aspdfdropdown = self.pyautogui.center(aspdfdropdown)
                x, y = aspdfdropdown
                self.pyautogui.click(x, y)
                self.aspdfdropdown = True
                break
        while self.emailaspdf == False:
            emailaspdf = self.pyautogui.locateOnScreen('emailaspdf.png', confidence=0.9)
            if emailaspdf != None:
                emailaspdf = self.pyautogui.center(emailaspdf)
                x, y = emailaspdf
                self.pyautogui.click(x, y)
                print('I have selected the Email Proforma as PDF option.')
                self.emailaspdf = True
                break
        while self.emailto == False:
            emailto = self.pyautogui.locateOnScreen('emailto.png', confidence=0.9)
            if emailto != None:
                emailto = self.pyautogui.center(emailto)
                x, y = emailto
                self.pyautogui.click(x, y)
                print('I have selected the Email To box.')
                self.emailto = True
                break
        while self.jobcontactemail == False:
            jobcontactemail = self.pyautogui.locateOnScreen('jobcontactemail.png', confidence=0.9)
            if jobcontactemail != None:
                jobcontactemail = self.pyautogui.center(jobcontactemail)
                x, y = jobcontactemail
                self.pyautogui.moveTo(x, y)
                self.pyautogui.click(clicks=2)
                print('I have selected the Main Job Contact email.')
                self.jobcontactemail = True
                break
        num = 0
        x = 6
        while num < x:
            try:
                email().selectallaccountsemails(num)
                num = num + 1
            except:
                num = num + 1
        print('I have checked for all accounts  emails and CCed them if found.')
        emailto_ok = self.pyautogui.locateOnScreen('emailto_ok.png', confidence=0.9)
        emailto_ok = self.pyautogui.center(emailto_ok)
        x, y = emailto_ok
        self.pyautogui.click(x, y)
        while self.displayasemail == False:
            displayasemail = self.pyautogui.locateOnScreen('displayasemail.png', confidence=0.9)
            if displayasemail != None:
                displayasemail = self.pyautogui.center(displayasemail)
                x, y = displayasemail
                self.pyautogui.click(x, y)
                print('I have selected the display as email tickbox.')
                self.displayasemail = True
                break
        emailsubject = self.pyautogui.locateOnScreen('emailsubject.png', confidence=0.9)
        emailsubject = self.pyautogui.center(emailsubject)
        x, y = emailsubject
        self.pyautogui.click(x, y)
        self.pyautogui.press('space')
        self.pyautogui.typewrite(self.jobref)

        emailsend = self.pyautogui.locateOnScreen('emailsend.png', confidence=0.9)
        emailsend = self.pyautogui.center(emailsend)
        x, y = emailsend
        self.pyautogui.click(x, y)
        while self.popup == False:
            popup = self.pyautogui.locateOnScreen('allowaccessfor.png', confidence=0.9)
            if popup != None:
                popup = self.pyautogui.center(popup)
                x, y = popup
                self.pyautogui.click(x, y)
                allow = self.pyautogui.locateOnScreen('allow.png', confidence=0.9)
                allow = self.pyautogui.center(allow)
                x, y = allow
                self.pyautogui.click(x, y)
                print('I have selected the email-Allow access for 1 minute-popup.')
                self.popup = True
            else:
                print('Please wait for popup to load.')
                continue

        while self.proformaemailwindow == False:
            proformaemailwindow = self.pyautogui.locateOnScreen('proformaemailwindow.png', confidence=0.7)
            if proformaemailwindow != None:
                print('Proforma email window has opened')
                attachfile = self.pyautogui.locateOnScreen('attachfile.png', confidence=0.9)
                attachfile = self.pyautogui.center(attachfile)
                x, y = attachfile
                self.pyautogui.click(x, y)
                self.proformaemailwindow = True
        while self.TCs == False:
            TCs = self.pyautogui.locateOnScreen('TCs.png', confidence=0.9)
            if TCs != None:
                TCs = self.pyautogui.locateOnScreen('TCs.png', confidence=0.9)
                TCs = self.pyautogui.center(TCs)
                x, y = TCs
                self.pyautogui.click(x, y)
                print('I have now attached Terms & Conditions file.')
                self.TCs = True


    def selectallaccountsemails(self, num):

        def ACCOUNT_DIRECTOR():
            ACCOUNTDIRECTOR = self.pyautogui.locateOnScreen('ACCOUNTDIRECTOR.png', confidence=0.9)
            ACCOUNTDIRECTOR = self.pyautogui.center(ACCOUNTDIRECTOR)
            x, y = ACCOUNTDIRECTOR
            self.pyautogui.moveTo(x, y)
            self.pyautogui.click(clicks=2)

        def ACCOUNT_MANAGER():
            ACCOUNTMANAGER = self.pyautogui.locateOnScreen('ACCOUNTMANAGER.png', confidence=0.9)
            ACCOUNTMANAGER = self.pyautogui.center(ACCOUNTMANAGER)
            x, y = ACCOUNTMANAGER
            self.pyautogui.moveTo(x, y)
            self.pyautogui.click(clicks=2)

        def ACCOUNTANT_():
            ACCOUNTANT = self.pyautogui.locateOnScreen('ACCOUNTANT.png', confidence=0.9)
            ACCOUNTANT = self.pyautogui.center(ACCOUNTANT)
            x, y = ACCOUNTANT
            self.pyautogui.moveTo(x, y)
            self.pyautogui.click(clicks=2)

        def ppurchasing_email():
            ppurchasingemail = self.pyautogui.locateOnScreen('ppurchasingemail.png', confidence=0.9)
            ppurchasingemail = self.pyautogui.center(ppurchasingemail)
            x, y = ppurchasingemail
            self.pyautogui.moveTo(x, y)
            self.pyautogui.click(clicks=2)

        def PURCHASE_LEDGER():
            PURCHASELEDGER = self.pyautogui.locateOnScreen('PURCHASELEDGER.png', confidence=0.9)
            PURCHASELEDGER = self.pyautogui.center(PURCHASELEDGER)
            x, y = PURCHASELEDGER
            self.pyautogui.moveTo(x, y)
            self.pyautogui.click(clicks=2)

        mydict = {'ACCOUNT_DIRECTOR': ACCOUNT_DIRECTOR,
                  'ACCOUNT_MANAGER': ACCOUNT_MANAGER,
                  'ACCOUNTANT_': ACCOUNTANT_,
                  'ppurchasing_email': ppurchasing_email,
                  'PURCHASE_LEDGER': PURCHASE_LEDGER}
        keynames = ['ACCOUNT_DIRECTOR', 'ACCOUNT_MANAGER', 'ACCOUNTANT_', 'ppurchasing_email', 'PURCHASE_LEDGER']

        return mydict[keynames[num]]()

    # def r_vat(self):
    #
    # def r_addinv(self):
    #
    # def r_baddebt(self):
    #
    # def miai(self):
    #
    # def r_receipt(self, invoicenumber, jobref, amountpaid):
    # if jobref != 0:
    # if invoicenumber == 0:
    # then open job by jobref
    # elif jobref == 0:
    # if invoicenumber != 0:
    # open job by invoicenumber
    # else:
    # print(please enter either 'invoicenumber, 0, amountpaid')
    # print(or enter '0, jobref, amountpaid')

    # if exact amount of invoice is recorded on the job:
    # then send receipt email


class select:
    import pyautogui
    import time
    import os
    import pyperclip
    import keyboard

    def __init__(self):
        self.os.chdir(r'C:\Users\Aimee\Pictures\TSS Python Screenshots\proforma reminder')
        self.allhiretabunsel = False
        self.allhiretabsel = False
        self.findref = False
        self.searchref = False
        self.jobfound = False
        self.jobopened = False

    def allhiretab(self):
        while self.allhiretabsel == False:
            allhiretabunsel = self.pyautogui.locateOnScreen('allhiretabunsel.png', confidence=0.9)
            allhiretabsel = self.pyautogui.locateOnScreen('allhiretabsel.png', confidence=0.9)
            if allhiretabunsel != None:
                allhiretabunsel = self.pyautogui.center(allhiretabunsel)
                x, y = allhiretabunsel
                self.pyautogui.click(x, y)
                print('I have selected the allhire tab')
                self.allhiretabunsel = False
                self.allhiretabsel = True
                break
            elif allhiretabsel != None:
                print('allhire tab is already selected')
                self.allhiretabunsel = False
                self.allhiretabsel = True
                break
            else:
                print('Cannot see all hire tab at all')
                allhiretabunsel = self.pyautogui.locateOnScreen('allhiretabunsel.png', confidence=0.9)
                allhiretabsel = self.pyautogui.locateOnScreen('allhiretabsel.png', confidence=0.9)
                TSSicon = self.pyautogui.locateOnScreen('TSSicon.png', confidence=0.9)
                TSSicon = self.pyautogui.center(TSSicon)
                x, y = TSSicon
                self.pyautogui.click(x, y)
                print('I can now see it!')
                if allhiretabunsel != None:
                    allhiretabunsel = self.pyautogui.center(allhiretabunsel)
                    x, y = allhiretabunsel
                    self.pyautogui.click(x, y)
                    self.allhiretabunsel = False
                    self.allhiretabsel = True
                    break
                elif allhiretabsel != None:
                    print('allhire tab is already selected')
                    self.allhiretabunsel = False
                    self.allhiretabsel = True
                    break

    def openjob(self, jobref):
        while self.findref == False:
            findref = self.pyautogui.locateOnScreen('findref.png', confidence=0.9)
            if findref != None:
                findref = self.pyautogui.center(findref)
                x, y = findref
                self.pyautogui.click(x, y)
                print('I have found the Find reference search box!')
                self.findref = True
                break
            else:
                print('I cannot locate the Find reference search box')
                TSSallhirescreen = self.pyautogui.locateOnScreen('TSSallhirescreen.png', confidence=0.9)
                if TSSallhirescreen != None:
                    TSSallhirescreen = self.pyautogui.center(TSSallhirescreen)
                    x, y = TSSallhirescreen
                    self.pyautogui.click(x, y)
                    print('I have now selected the TSSallhirescreen')
                    select().allhiretab()
                else:
                    TSSicon = self.pyautogui.locateOnScreen('TSSicon.png', confidence=0.9)
                    TSSicon = self.pyautogui.center(TSSicon)
                    x, y = TSSicon
                    self.pyautogui.click(x, y)
                    select().allhiretab()
                continue

        while self.searchref == False:
            highlightref = self.pyautogui.locateOnScreen('highlightref.png', confidence=0.9)
            if highlightref != None:
                highlightref = self.pyautogui.center(highlightref)
                x, y = highlightref
                print('I will now enter the job reference number {}'.format(jobref))
                self.pyautogui.dragTo(x, y, button='left')
                self.pyautogui.press('backspace')
                self.pyautogui.typewrite(str(jobref))
                print('I have now entered the job reference number {}'.format(jobref))
                searchref = self.pyautogui.locateOnScreen('searchref.png', confidence=0.9)
                searchref = self.pyautogui.center(searchref)
                x, y = searchref
                self.pyautogui.click(x, y)
                print('I have clicked on the search button. Lets now wait for job {} to load'.format(jobref))
                while self.jobfound == False:
                    jobfound = self.pyautogui.locateOnScreen('jobfound.png', confidence=0.9)
                    if jobfound != None:
                        jobfound = self.pyautogui.center(jobfound)
                        x, y = jobfound
                        self.pyautogui.moveTo(x, y)
                        self.pyautogui.click(clicks=2)
                        self.jobfound = True
                        print(
                            'Job {} has now loaded and have double-clicked on it. Lets now wait for it to open'.format(
                                jobref))
                        while self.jobopened == False:
                            readonly = self.pyautogui.locateOnScreen('readonly.png', confidence=0.9)
                            jobopened = self.pyautogui.locateOnScreen('jobopened.png', confidence=0.9)
                            if readonly != None:
                                print('Job ref {} has opened, but someone else is on the job!'.format(jobref))
                                closejob = self.pyautogui.locateOnScreen('closejob.png', confidence=0.9)
                                x = closejob.left + closejob.width - 10
                                y = closejob.top + 20
                                self.pyautogui.click(x, y)
                                print('I have now closed job ref {}.'.format(jobref))
                                break
                            elif jobopened != None:
                                self.jobopened = True
                                print('Job {} has opened!'.format(jobref))
                                break
                            else:
                                self.time.sleep(2)
                                print('Job {} has not opened yet. Please wait.'.format(jobref))
                                continue
                        break
                    else:
                        print('Job {} has not yet loaded. Please wait.'.format(jobref))
                        continue
                break
            else:
                print('I cannot highlight the job reference because I cannot locate the highlightref area')
                select().allhiretab()
                continue


class reminder:
    import pyautogui
    import time
    import os
    import pyperclip
    import keyboard

    def __init__(self):
        self.os.chdir(r'C:\Users\Aimee\Pictures\TSS Python Screenshots\proforma reminder')

    def sendreminder(self):
        select().allhiretab()

# reminder().sendreminder()
# select().openjob('069339')
email().r_proforma()