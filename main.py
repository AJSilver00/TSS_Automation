class email:
    import pyautogui
    import time
    import os
    import pyperclip
    import keyboard

    def __init__(self):
        self.os.chdir(r'C:\Users\Aimee\Pictures\TSS Python Screenshots\proforma reminder')
        self.invoicingtabunsel = False
        self.invoicingtabsel = False
        self.invoiceproductionsel = False
#    def r2(self):
        # if paymentduedate is before startdate:
            # print('This is not an acc cust.')
            # if todaysdate is 1 week b4 startdate then send 'proforma1' email
            # elif startdate is tomorrow then send 'proforma2' email AND make a note for myself to ring them up telling them to pay

        # else:
            # print('This is an acc cust.')
            # if todaysdate is less than 1 week b4 paymentduedate and remindersentdate is 1 week ago then do nothing
            # if todaysdate is less than 1 week b4 paymentduedate and remindersentdate is more than 1 week ago then send 'vat1' email.
            # elif todaysdate is more than 1 week after paymentduedate and remindersentdate is null then print('no reminder sent yet! They are overdue!') and send 'vat3' email. Also make sure remindersentdate is not null anymore once the email is sent.
            # elif todaysdate is more than 1 week after paymentduedate and remindersentdate is 1 or more weeks ago then print('no reminder sent yet!' and make a note of it for myself to check this later, as I should have sent one.)
            # else:
                # if todaysdate is exactly 1 week b4 paymentduedate then send 'vat1' email
                    # if remindersentdate is past paymentduedate by more than 1 week, then make a note of it for myself to check this later, as it should have been sent exactly 1 week past paymentduedate!
                # elif todaysdate == paymentduedate then send 'vat2' email
                    # if remindersentdate is null OR is more than 1 week b4 paymentduedate, then make a note of it for myself to check this later, as it should have been sent exactly 1 week before paymentduedate!
                # else:
                    # if todaysdate is past paymentduedate by 1 week
                        # elif remindersentdate was exactly 1 week ago then send 'vat3' email.
                    # if number of reminders sent past paymentduedate is 3:
                        # then send 'baddebt2' email
                    # elif number of reminders sent past paymentduedate is > 3:
                        # then make a note of it, so I can check this later.
                    # else just send 'vat3' email.
                # else go onto next row in Excel report and continue loop (i.e. restart the loop).

    def r_proforma(self):
        while self.invoicingtabsel == False:
            invoicingtabunsel = self.pyautogui.locateOnScreen('invoicingtabunsel.png', confidence=0.99)
            invoicingtabsel = self.pyautogui.locateOnScreen('invoicingtabsel.png', confidence=0.99)
            if invoicingtabunsel != None:
                invoicingtabunsel = self.pyautogui.center(invoicingtabunsel)
                x, y = invoicingtabunsel
                self.pyautogui.click(x, y)
                print('I have selected the Invoicing tab')
                self.invoicingtabsel = True
            else:
                invoicingtabsel = self.pyautogui.locateOnScreen('invoicingtabsel.png',  confidence=0.99)
                if invoicingtabsel != None:
                    print('Invoicing tab is already selected')
                    self.invoicingtabsel = True
                    break
                else:
                    print('Cannot find Invoice Production tab yet. Please wait.')
        while self.invoiceproductionsel == False:
            invoiceproductionunsel = self.pyautogui.locateOnScreen('invoiceproductionunsel.png', confidence=0.99)
            invoiceproductionsel = self.pyautogui.locateOnScreen('invoiceproductionsel.png', confidence=0.99)
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
                        print('Job {} has now loaded and have double-clicked on it. Lets now wait for it to open'.format(jobref))
                        while self.jobopened == False:
                            readonly = self.pyautogui.locateOnScreen('readonly.png', confidence=0.9)
                            jobopened = self.pyautogui.locateOnScreen('jobopened.png', confidence=0.9)
                            if readonly != None:
                                print('Job ref {} has opened, but someone else is on the job!'.format(jobref))
                                closejob = self.pyautogui.locateOnScreen('closejob.png', confidence=0.9)
                                x = closejob.left + closejob.width - 10
                                y = closejob.top + 20
                                self.pyautogui.click(x,y)
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

#reminder().sendreminder()
#select().openjob('069339')
email().r_proforma()