class MouseAutomation:
    import pyautogui
    import time
    import os
    import pyperclip
    pyautogui.PAUSE = 1

    def __init__(self):
        self.os.chdir(r'C:\Users\Aimee\Pictures\TSS Python Screenshots')


    def search_JobRef(self, jobref):
        FindRef = self.pyautogui.locateOnScreen('FindRef.png', confidence=0.9)
        if FindRef != None:
            FindRef = self.pyautogui.center(FindRef)
            x, y = FindRef
            self.pyautogui.click(x, y)
        DragTO = self.pyautogui.locateOnScreen('FindRef dragleft.png', confidence=0.9)
        if DragTO != None:
            DragTO = self.pyautogui.center(DragTO)
            x, y = DragTO
            self.pyautogui.dragTo(x, y, button='left')
            self.pyautogui.press('backspace')  # deletes what is entered in box
            self.pyautogui.typewrite(str(jobref))
        SearchRef = self.pyautogui.locateOnScreen('SearchRef.png', confidence=0.9)
        if SearchRef != None:
            SearchRef = self.pyautogui.center(SearchRef)
            x, y = SearchRef
            self.pyautogui.click(x, y)
            self.time.sleep(3)

    def openjob(self, jobref):  # function opens a job
        mouse.click_allhiretab()
        mouse.search_JobRef(jobref)
        doubleclickjob = self.pyautogui.locateOnScreen('jobsearched.png', confidence=0.9)
        if doubleclickjob != None:
            doubleclickjob = self.pyautogui.center(doubleclickjob)
            x, y = doubleclickjob
            self.pyautogui.moveTo(x, y)
            self.pyautogui.click(clicks=2)
            self.time.sleep(60)  # waits 60sec for job to open
        tab = self.pyautogui.locateOnScreen('accountonstoptab.png', confidence=0.9)
        okbutton = self.pyautogui.locateOnScreen('accountonstopOKbutton.png', confidence=0.9)
        if tab != None:
            if okbutton != None:
                okbutton = self.pyautogui.center(okbutton)
                x, y = okbutton
                self.pyautogui.click(x, y)
                self.time.sleep(3)
        readonly = self.pyautogui.locateOnScreen('readonly.png', confidence=0.9)
        if readonly != None:
            closejob = self.pyautogui.locateOnScreen('jobtabXbutton.png', confidence=0.9)
            closejob = self.pyautogui.center(closejob)
            x, y = closejob
            self.pyautogui.click(x, y)
            print('ERROR: Someone else is on the Job')

    def excelfindmacro(self):
        self.pyautogui.PAUSE = 2
        excelFile = self.pyautogui.locateOnScreen('excelFile.png', confidence=0.9)
        excelFile = self.pyautogui.center(excelFile)
        x, y = excelFile
        self.pyautogui.click(x, y)

        excelFileOpen = self.pyautogui.locateOnScreen('excelFileOpen.png', confidence=0.9)
        excelFileOpen = self.pyautogui.center(excelFileOpen)
        x, y = excelFileOpen
        self.pyautogui.click(x, y)

        reportcsvOpen = self.pyautogui.locateOnScreen('reportcsvOpen.png', confidence=0.9)
        reportcsvOpen = self.pyautogui.center(reportcsvOpen)
        x, y = reportcsvOpen
        self.pyautogui.click(x, y)
        self.time.sleep(6)

        maximizeExcel = self.pyautogui.locateOnScreen('maximizeExcel.png', confidence=0.9)
        maximizeExcel = self.pyautogui.center(maximizeExcel)
        x, y = maximizeExcel
        self.pyautogui.click(x, y)

        excelviewtab = self.pyautogui.locateOnScreen('excelviewtab.png', confidence=0.9)
        excelviewtab = self.pyautogui.center(excelviewtab)
        x, y = excelviewtab
        self.pyautogui.click(x, y)

        macrodropdown = self.pyautogui.locateOnScreen('macrodropdown.png', confidence=0.9)
        macrodropdown = self.pyautogui.center(macrodropdown)
        x, y = macrodropdown
        self.pyautogui.click(x, y)

        viewmacros = self.pyautogui.locateOnScreen('viewmacros.png', confidence=0.9)
        viewmacros = self.pyautogui.center(viewmacros)
        x, y = viewmacros
        self.pyautogui.click(x, y)

    def runexcelreport(self):
        self.pyautogui.PAUSE = 2
        mouse.excelfindmacro()
        macroReportReformatter = self.pyautogui.locateOnScreen('macroReportReformatter.png', confidence=0.9)
        macroReportReformatter = self.pyautogui.center(macroReportReformatter)
        x, y = macroReportReformatter
        self.pyautogui.click(x, y)

        macroReportReformatterRun = self.pyautogui.locateOnScreen('macroReportReformatterRun.png', confidence=0.9)
        macroReportReformatterRun = self.pyautogui.center(macroReportReformatterRun)
        i, j = macroReportReformatterRun
        self.pyautogui.click(i, j)

        macrodropdown = self.pyautogui.locateOnScreen('macrodropdown.png', confidence=0.9)
        macrodropdown = self.pyautogui.center(macrodropdown)
        x, y = macrodropdown
        self.pyautogui.click(x, y)

        viewmacros = self.pyautogui.locateOnScreen('viewmacros.png', confidence=0.9)
        viewmacros = self.pyautogui.center(viewmacros)
        x, y = viewmacros
        self.pyautogui.click(x, y)

        macroSendEmails = self.pyautogui.locateOnScreen('macroSendEmails.png', confidence=0.9)
        macroSendEmails = self.pyautogui.center(macroSendEmails)
        x, y = macroSendEmails
        self.pyautogui.click(x, y)
        self.pyautogui.click(i, j)

    def sendReminder(self, jobref):
        mouse.openjob(jobref)
        self.pyautogui.PAUSE = 4

        detailsTab = self.pyautogui.locateOnScreen('detailsTab.png', confidence=0.9)
        detailsTab = self.pyautogui.center(detailsTab)
        x, y = detailsTab
        self.pyautogui.click(x, y)
        self.pyautogui.hotkey('ctrl', 'c')

        invoicingTab = self.pyautogui.locateOnScreen('invoicingTab.png', confidence=0.9)
        invoicingTab = self.pyautogui.center(invoicingTab)
        x, y = invoicingTab
        self.pyautogui.click(x, y)

        previouslyRaisedInvoicesTab = self.pyautogui.locateOnScreen('previouslyRaisedInvoicesTab.png', confidence=0.9)
        previouslyRaisedInvoicesTab = self.pyautogui.center(previouslyRaisedInvoicesTab)
        x, y = previouslyRaisedInvoicesTab
        self.pyautogui.click(x, y)

        mainHireInvRaised = self.pyautogui.locateOnScreen('mainHireInvRaised.png', confidence=0.9)
        mainHireInvRaised = self.pyautogui.center(mainHireInvRaised)
        x, y = mainHireInvRaised
        self.pyautogui.click(x, y)

        EmailInvButton = self.pyautogui.locateOnScreen('EmailInvButton.png', confidence=0.9)
        EmailInvButton = self.pyautogui.center(EmailInvButton)
        x, y = EmailInvButton
        self.pyautogui.click(x, y)

        emailPDFButton = self.pyautogui.locateOnScreen('emailPDFButton.png', confidence=0.9)
        emailPDFButton = self.pyautogui.center(emailPDFButton)
        x, y = emailPDFButton
        self.pyautogui.click(x, y)

        AttachmentInvLoaded = self.pyautogui.locateOnScreen('AttachmentInvLoaded.png', confidence=0.9)
        if AttachmentInvLoaded != None:
            EmailCompleteButton = self.pyautogui.locateOnScreen('EmailCompleteButton.png', confidence=0.9)
            EmailCompleteButton = self.pyautogui.center(EmailCompleteButton)
            x, y = EmailCompleteButton
            self.pyautogui.click(x, y)

            EmailTo = self.pyautogui.locateOnScreen('EmailTo.png', confidence=0.9)
            EmailTo = self.pyautogui.center(EmailTo)
            x, y = EmailTo
            self.pyautogui.click(x, y)

            JOBCONTACTEmail = self.pyautogui.locateOnScreen('JOBCONTACTEmail.png', confidence=0.9)
            JOBCONTACTEmail = self.pyautogui.center(JOBCONTACTEmail)
            x, y = JOBCONTACTEmail
            self.pyautogui.moveTo(x, y)
            self.pyautogui.click(clicks=2)

            ACCOUNTANTEmail = self.pyautogui.locateOnScreen('ACCOUNTANTEmail.png', confidence=0.9)
            if ACCOUNTANTEmail != None:
                ACCOUNTANTEmail = self.pyautogui.center(ACCOUNTANTEmail)
                x, y = ACCOUNTANTEmail
                self.pyautogui.moveTo(x, y)
                self.pyautogui.click(clicks=2)

            EmailRecipientOKButton = self.pyautogui.locateOnScreen('EmailRecipientOKButton.png', confidence=0.9)
            EmailRecipientOKButton = self.pyautogui.center(EmailRecipientOKButton)
            x, y = EmailRecipientOKButton
            self.pyautogui.click(x, y)

            EmailSubject = self.pyautogui.locateOnScreen('EmailSubject.png', confidence=0.9)
            EmailSubject = self.pyautogui.center(EmailSubject)
            x, y = EmailSubject
            self.pyautogui.click(x, y)
            self.pyautogui.press('space')

            sendAsOutlookEmailTickButton = self.pyautogui.locateOnScreen('sendAsOutlookEmailTickButton.png', confidence=0.9)
            sendAsOutlookEmailTickButton = self.pyautogui.center(sendAsOutlookEmailTickButton)
            x, y = sendAsOutlookEmailTickButton
            self.pyautogui.click(x, y)

            sendEmailButton = self.pyautogui.locateOnScreen('sendEmailButton.png', confidence=0.9)
            sendEmailButton = self.pyautogui.center(sendEmailButton)
            x, y = sendEmailButton
            self.pyautogui.click(x, y)
            self.time.sleep(10)

            outlookFlashing = self.pyautogui.locateOnScreen('outlookFlashing.png', confidence=0.9)
            if outlookFlashing != None:
                outlookFlashing = self.pyautogui.center(outlookFlashing)
                x, y = outlookFlashing
                self.pyautogui.click(x, y)

                outlookFlashing1 = self.pyautogui.locateOnScreen('outlookFlashing1.png', confidence=0.9)
                outlookFlashing1 = self.pyautogui.center(outlookFlashing1)
                x, y = outlookFlashing1
                self.pyautogui.click(x, y)

            POPUP = self.pyautogui.locateOnScreen('POPUP.png', confidence=0.9)
            if POPUP != None:
                POPUP = self.pyautogui.center(POPUP)
                x, y = POPUP
                self.pyautogui.click(x, y)

                POPUPTick = self.pyautogui.locateOnScreen('POPUPTick.png', confidence=0.9)
                POPUPTick = self.pyautogui.center(POPUPTick)
                x, y = POPUPTick
                self.pyautogui.click(x, y)

                POPUPAllow = self.pyautogui.locateOnScreen('POPUPAllow.png', confidence=0.9)
                POPUPAllow = self.pyautogui.center(POPUPAllow)
                x, y = POPUPAllow
                self.pyautogui.click(x, y)

    def attachInvToEmail(self):
        mouse.runexcelreport()
        self.pyautogui.PAUSE = 2
        excelminimize = self.pyautogui.locateOnScreen('excelminimize.png', confidence=0.9)
        excelminimize = self.pyautogui.center(excelminimize)
        x, y = excelminimize
        self.pyautogui.click(x, y)

        msgemail = self.pyautogui.locateOnScreen('msgemail.png', confidence=0.9)
        msgemail = self.pyautogui.center(msgemail)
        x, y = msgemail
        self.pyautogui.click(x, y)

        counter = 0
        emailRef = self.pyautogui.locateOnScreen('emailRef.png', confidence=0.9)
        while True:
            if emailRef == None:
                counter=counter+1
            else:
                mouse.highlightEmailRef()
                email_jobref = self.pyperclip.paste()
                mouse.sendReminder(email_jobref)

    def highlightEmailRef(self):
        self.pyautogui.PAUSE = 0.2
        RefSelect1 = self.pyautogui.locateOnScreen('RefSelect1.png', confidence=0.9)
        RefSelect1 = self.pyautogui.center(RefSelect1)
        a, b = RefSelect1

        emailRef = self.pyautogui.locateOnScreen('emailRef.png', confidence=0.9)
        edgeX = emailRef.left + emailRef.width
        edgeY = emailRef.top
        self.pyautogui.click(edgeX, edgeY)
        self.pyautogui.dragTo(a, b, button='left')
        self.pyautogui.hotkey('ctrl', 'c')

    # def sendReceipt(self, jobref, invNO):
    #     if jobref = 0:
    #     # send receipt by invNO
    #     elif invNO = 0:
    #     # send receipt by jobref
    #     else:
    #         print('No jobref or invNO given')

    # def sendReminder(self):
    #     click = self.clickhere = self.pyautogui.click(self.coor)
    #
    # def sendReportToExcel(self):
    #     click = self.click = self.pyautogui.click()
    #
    # def send(self):
    #     click = self.click = self.pyautogui.click()

mouse = MouseAutomation()
mouse.attachInvToEmail()
#mouse.test()