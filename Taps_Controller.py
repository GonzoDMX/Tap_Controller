#!/usr/bin/python
# -*- coding: utf-8 -*-

import wx
import os
import sys
import binascii
import time
import serial
import serial.tools.list_ports as port_list
from serial.serialutil import SerialException

import xlsxwriter

serOpen = False
selCom = ''
modeSelect = 0

gFade = 0

success = False

recGroup = "A"
sensGroup = 5

Disclaimer1 = "Tap Shoes Controller Created by: Andrew O\'Shei"
Disclaimer2 = "Do not use, copy or distribute without permission"
Disclaimer3 = "For more info contact: andrewoshei@gmail.com"

selectedCue = False
selCue = -1
cueCount = 0
cueTrig = False
setLink = False

cueContainer = ["", "1", "All", "Solid", "A", "0", "0,0,0", "0,0,0"]
nonitem = ", , , , , , "

class BuildGUI(wx.Frame):
           
    def __init__(self, *args, **kw):
        super(BuildGUI, self).__init__(*args, **kw)
        #self.timer = wx.Timer(self)
        #self.Bind(wx.EVT_TIMER, self.Set_Output, self.timer)
        #self.timer.Start(200)
        
        self.greenTimer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.coolGreen, self.greenTimer)
        
        self.Bind(wx.EVT_CLOSE, self.OnClose)
        icoF = 'taps.ico'
        iconS = wx.Icon(icoF, wx.BITMAP_TYPE_ICO)
        self.SetIcon(iconS)
        self.InitGUI()
        
        
    def InitGUI(self):
        global ComPorts
        self.glist = ["All", "A", "B"]
        self.modelist = ["Solid Color", "Tap React", "Effect"]
        self.shortList = ["Solid", "React", "FX"]
        self.fxNames = ["Rainbow", "Strobe", "RedStrobe", "BluStrobe", "GrnStrobe", "RedSpark", "Fade", "RedFade", "BluFade", "GrnFade", "Love", "Fire", "Ice", "RedWave", "BluWave", "GrnWave", "RedBounce", "BluBounce", "GrnBounce"]
        self.fxlist = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S"]
        self.collist = [" ", "Cue", "Target", "Mode", "FX", "S", "Color 1", "Color 2"]
        self.colwidth = [20, 35, 50, 45, 35, 35, 75, 75]
        self.pnl = wx.Panel(self)

        self.LabelSel = wx.StaticText(self.pnl, label='Select Device:', pos=(10, 13))
        self.cb = wx.ComboBox(self.pnl, size=(180, 22), pos=(90, 10), style=wx.CB_READONLY)
        self.cb.Bind(wx.EVT_COMBOBOX_DROPDOWN, self.Refresh_Dev_List)
        self.cb.Bind(wx.EVT_COMBOBOX, self.On_Dev_Select)


        self.rebu = wx.Button(self.pnl, label='Refresh', size=(50, 24), pos=(270, 9))
        self.rebu.Bind(wx.EVT_BUTTON, self.Refresh_Dev_List)

        self.LabelDev = wx.StaticText(self.pnl, label='Device Type:', pos=(10, 38))
        self.DevTyp = wx.TextCtrl(self.pnl, style=wx.TE_READONLY, size=(230, 22), pos=(90, 35))        
        self.DevTyp.Disable()
        
        self.LabelAdd = wx.StaticText(self.pnl, label='DMX Address:', pos=(10, 68))
        self.AddDMX = wx.SpinCtrl(self.pnl, min=1, max=512, size=(80, 22), pos=(90, 65))
        self.addbu = wx.Button(self.pnl, label='Set New Address', size=(100, 24), pos=(180, 64))
        self.addbu.Bind(wx.EVT_BUTTON, self.Update_Device)
        self.AddDMX.Disable()
        self.addbu.Disable()

        self.LabelSel2 = wx.StaticText(self.pnl, label='Target:', pos=(350, 33))
        self.cb2 = wx.ComboBox(self.pnl, value=self.glist[0], size=(50, 22), pos=(345, 55), style=wx.CB_READONLY, choices=self.glist)
        self.cb2.Bind(wx.EVT_COMBOBOX, self.On_Target_Select)
        
        self.LabelSel3 = wx.StaticText(self.pnl, label='Mode:', pos=(20, 108))
        self.cb3 = wx.ComboBox(self.pnl, value=self.modelist[0], size=(85, 20),
                               pos=(60, 105), style=wx.CB_READONLY, choices=self.modelist)
        self.cb3.Bind(wx.EVT_COMBOBOX, self.On_Mode_Select)
        
    
        """Controls for setting Color 1 Param."""
        self.bBox1 = wx.StaticBox(self.pnl, label="Color 1", size=(315, 120), pos=(10, 135))
        self.colorPrev1 = wx.TextCtrl(self.pnl, style=wx.TE_READONLY, pos=(95, 155), size=(150, 15))
        self.colorPrev1.SetBackgroundColour((0, 0, 0))
        self.R1Label = wx.StaticText(self.pnl, label='Red:', pos=(20, 178))
        self.R1Slider = wx.Slider(self.pnl, style=wx.SL_HORIZONTAL, minValue=0, maxValue=255, pos=(55, 175), size=(230,22))
        self.R1Slider.Bind(wx.EVT_SLIDER, self.Set_Color_1)
        self.R1Count = wx.TextCtrl(self.pnl, style=wx.TE_READONLY, value="0", size=(30, 22), pos=(285, 174)) 
        self.G1Label = wx.StaticText(self.pnl, label='Green:', pos=(20, 203))
        self.G1Slider = wx.Slider(self.pnl, style=wx.SL_HORIZONTAL, minValue=0, maxValue=255, pos=(55, 200), size=(230,22))
        self.G1Slider.Bind(wx.EVT_SLIDER, self.Set_Color_1)
        self.G1Count = wx.TextCtrl(self.pnl, style=wx.TE_READONLY, value="0", size=(30, 22), pos=(285, 199))
        self.B1Label = wx.StaticText(self.pnl, label='Blue:', pos=(20, 228))
        self.B1Slider = wx.Slider(self.pnl, style=wx.SL_HORIZONTAL, minValue=0, maxValue=255, pos=(55, 225), size=(230,22))
        self.B1Slider.Bind(wx.EVT_SLIDER, self.Set_Color_1)
        self.B1Count = wx.TextCtrl(self.pnl, style=wx.TE_READONLY, value="0", size=(30, 22), pos=(285, 224))


        
        """Controls for setting Color 2 Param."""
        self.bBox2 = wx.StaticBox(self.pnl, label="Color2", size=(315, 120), pos=(10, 255))
        self.colorPrev2 = wx.TextCtrl(self.pnl, style=wx.TE_READONLY, pos=(95, 275), size=(150, 15))
        self.colorPrev2.SetBackgroundColour((0, 0, 0))
        self.R2Label = wx.StaticText(self.pnl, label='Red:', pos=(20, 298))
        self.R2Slider = wx.Slider(self.pnl, style=wx.SL_HORIZONTAL, minValue=0, maxValue=255, pos=(55, 295), size=(230,22))
        self.R2Slider.Bind(wx.EVT_SLIDER, self.Set_Color_2)
        self.R2Count = wx.TextCtrl(self.pnl, style=wx.TE_READONLY, value="0", size=(30, 22), pos=(285, 294)) 
        self.G2Label = wx.StaticText(self.pnl, label='Green:', pos=(20, 323))
        self.G2Slider = wx.Slider(self.pnl, style=wx.SL_HORIZONTAL, minValue=0, maxValue=255, pos=(55, 320), size=(230,22))
        self.G2Slider.Bind(wx.EVT_SLIDER, self.Set_Color_2)
        self.G2Count = wx.TextCtrl(self.pnl, style=wx.TE_READONLY, value="0", size=(30, 22), pos=(285, 319))
        self.B2Label = wx.StaticText(self.pnl, label='Blue:', pos=(20, 348))
        self.B2Slider = wx.Slider(self.pnl, style=wx.SL_HORIZONTAL, minValue=0, maxValue=255, pos=(55, 345), size=(230,22))
        self.B2Slider.Bind(wx.EVT_SLIDER, self.Set_Color_2)
        self.B2Count = wx.TextCtrl(self.pnl, style=wx.TE_READONLY, value="0", size=(30, 22), pos=(285, 344))
        
        self.bBox3 = wx.StaticBox(self.pnl, size=(315, 45), pos=(10, 90))
        
        """Control for setting the Effects Param."""
        self.LabelFX = wx.StaticText(self.pnl, label='FX:', pos=(155, 108))

        self.AddFX = wx.ComboBox(self.pnl, value=self.fxlist[0], size=(40, 20),
                               pos=(180, 105), style=wx.CB_READONLY, choices=self.fxlist)
        self.AddFX.Bind(wx.EVT_COMBOBOX, self.Set_Effect)
        
        """Control for setting the Speed Param."""
        self.LabelSpd = wx.StaticText(self.pnl, label='Speed:', pos=(230, 108))

        self.AddSpd = wx.SpinCtrl(self.pnl, min=0, max=10, initial=0 , size=(50, 20),
                               pos=(270, 105))
        self.AddSpd.Bind(wx.EVT_SPINCTRL, self.Set_Speed)
        

        self.st1 = wx.StaticText(self.pnl, label=Disclaimer1, pos=(40, 380))
        self.st2 = wx.StaticText(self.pnl, label=Disclaimer2, pos=(35, 395))
        self.st3 = wx.StaticText(self.pnl, label=Disclaimer3, pos=(38, 410))

        listStyle = wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.LC_HRULES
        
        self.cueList = wx.ListCtrl(self.pnl, style=listStyle, pos=(335, 98), size=(375, 270))
        for i in range(0, 8):
            self.cueList.InsertColumn(i, self.collist[i], width=self.colwidth[i])            

        self.cueList.Bind(wx.EVT_LIST_ITEM_SELECTED, self.List_Item_Selected)
        self.cueList.Bind(wx.EVT_LIST_ITEM_DESELECTED, self.List_Item_DeSelected)

        self.bBox4 = wx.StaticBox(self.pnl, label="Cue List", size=(385, 367), pos=(330, 8))
        
        self.cueBu1 = wx.Button(self.pnl, label='Add', pos=(420, 28), size=(70, 25))
        self.cueBu1.Bind(wx.EVT_BUTTON, self.Add_New_Cue)
        self.cueBu2 = wx.Button(self.pnl, label='Insert', pos=(500, 28), size=(70, 25))
        self.cueBu2.Bind(wx.EVT_BUTTON, self.Insert_Cue)
        self.cueBu3 = wx.Button(self.pnl, label='Link', pos=(580, 28), size=(70, 25))
        self.cueBu3.Bind(wx.EVT_BUTTON, self.Link_Cue)
        self.cueBu4 = wx.Button(self.pnl, label='Delete', pos=(420, 58), size=(70, 25))
        self.cueBu4.Bind(wx.EVT_BUTTON, self.Del_Cue)
        self.cueBu5 = wx.Button(self.pnl, label='Replace', pos=(500, 58), size=(70, 25))
        self.cueBu5.Bind(wx.EVT_BUTTON, self.Rep_Cue)
        self.cueBu6 = wx.Button(self.pnl, label='Clear All', pos=(580, 58), size=(70, 25))
        self.cueBu6.Bind(wx.EVT_BUTTON, self.Rem_All_Cues)
        
        self.cuePrint = wx.Button(self.pnl, label='Print', pos=(655, 42), size=(50, 25))
        self.cuePrint.Bind(wx.EVT_BUTTON, self.test_Save_File)      

        self.cueBu7 = wx.Button(self.pnl, label='Test', pos=(360, 390))
        self.cueBu7.Bind(wx.EVT_BUTTON, self.Test_Cue_Send)
        self.cueBu8 = wx.Button(self.pnl, label='GO', pos=(470, 390))
        self.cueBu8.Bind(wx.EVT_BUTTON, self.Go_Cue)        
        self.cueBu9 = wx.Button(self.pnl, label='STOP', pos=(580, 390))
        self.cueBu9.Bind(wx.EVT_BUTTON, self.Stop_Cue)

        self.cb2.Disable()
        self.cb3.Disable()
        self.cueBu1.Disable()
        self.cueBu2.Disable()
        self.cueBu3.Disable()
        # self.cueBu4.Disable()
        self.cueBu5.Disable()
        # self.cueBu6.Disable()
        self.cueBu7.Disable()
        self.cueBu8.Disable()
        self.cueBu9.Disable()
        self.colorPrev1.Disable()
        self.R1Slider.Disable()
        self.G1Slider.Disable()
        self.B1Slider.Disable()
        self.colorPrev2.Disable()
        self.R2Slider.Disable()
        self.G2Slider.Disable()
        self.B2Slider.Disable()
        self.AddFX.Disable()
        self.AddSpd.Disable()

        
        self.SetSize((740, 470))
        self.SetMinSize((740, 470))
        self.SetMaxSize((740, 470))
        self.SetTitle('Tap Shoes Controller')
        self.Centre()
        self.Refresh_Dev_List(0)
        self.pnl.SetDoubleBuffered(True)
        self.Show(True)


    def List_Item_Selected(self, item):
        global selectedCue
        global selCue
        global nonitem
        selCue = item.GetIndex()
        if selCue == 0:
            itemx = ", ".join([self.cueList.GetItem(selCue, col).GetText()
                               for col in range(8)])
            if itemx == nonitem:
                self.cueList.Select(selCue, 0)
            else:
                selectedCue = True
                self.cueBu2.Enable()
                self.cueBu5.Enable()
        else:
            selectedCue = True
            self.cueBu2.Enable()
            self.cueBu5.Enable()


    def List_Item_DeSelected(self, item):
        global selectedCue
        global selCue
        selectedCue = False
        selCue = -1
        self.cueBu2.Disable()
        self.cueBu5.Disable()


    def Go_Cue(self, cue):
        global cueCount
        global cueTrig
        global selectedCue
        global selCue
        global setLink
        header = "<555"
        footer = ">"
        cueLim = self.cueList.GetItemCount()
        if selectedCue == True:
            if cueCount > 0:
                self.cueList.SetItem((cueCount - 1), 0, " ")
            self.cueList.SetItem(selCue, 0, ">")
            cueCount = selCue + 1
            self.cueList.Select(selCue, 0)
            cueTrig = True
            selectedCue = False
        else:
            if cueCount == cueLim:
                self.cueList.SetItem((cueCount - 1), 0, " ")
                self.cueList.SetItem(0, 0, ">")
                cueCount = 1
            else:
                if cueTrig == True and cueCount > 0:
                    self.cueList.SetItem((cueCount - 1), 0, " ")             
                    self.cueList.SetItem(cueCount, 0, ">")
                else:
                    self.cueList.SetItem(cueCount, 0, ">")
                cueCount += 1
                cueTrig = True
                
        targ = str(self.glist.index(self.cueList.GetItemText(cueCount - 1, 2))) 
        mod = str(self.shortList.index(self.cueList.GetItemText(cueCount - 1, 3)))
        fx = self.cueList.GetItemText(cueCount - 1, 4)
        spd = chr((int(self.cueList.GetItemText(cueCount - 1, 5)) + 48))
        col1 = self.cueList.GetItemText(cueCount - 1, 6).replace(",", " ").split()
        col2 = self.cueList.GetItemText(cueCount - 1, 7).replace(",", " ").split()
        
        color1 = ""
        color2 = ""
        for i in range(0, 3):
            a = str(int(((int(col1[i]) / 100) % 10)))
            b = str(int(((int(col1[i]) / 10) % 10)))
            c = str(int((int(col1[i]) % 10)))
            color1 = color1 + a + b + c
            a = str(int(((int(col2[i]) / 100) % 10)))
            b = str(int(((int(col2[i]) / 10) % 10)))
            c = str(int((int(col2[i]) % 10)))
            color2 = color2 + a + b + c

        if mod == "1":
            preMess = header + targ + mod + fx + spd  + color2 + footer
            message = preMess.encode('utf-8')
        else:
            preMess = header + targ + mod + fx + spd  + color1 + footer
            message = preMess.encode('utf-8')
        
        if cueCount < cueLim:
            coLink = self.cueList.GetItemBackgroundColour(cueCount)
            if coLink == (220, 220, 220, 255):
                setLink = True
        self.Send_Cue(message)
            

    def Send_Cue(self, message):
        global setLink
        """Send Test Cue over Serial"""
        print(message)
        ser = serial.Serial(selCom, 9600, timeout=0.5, write_timeout=0.5)
        try:
            ser.isOpen()
            ser.write(message)       
            ser.close()
        except SerialException:
            ser.close()
            self.DevTyp.Enable()
            self.DevTyp.SetForegroundColour((255, 255, 255))
            self.DevTyp.SetBackgroundColour((255, 0, 0))
            self.DevTyp.Clear()
            self.DevTyp.SetValue("Port Busy or Not Responding")
        if setLink == True:
            setLink = False
            time.sleep(0.2)
            self.Go_Cue("again")

    def Stop_Cue(self, cue):
        global cueCount
        global cueTrig
        global selectedCue
        global selCue
        if cueTrig == True:
            self.cueList.SetItem((cueCount - 1), 0, " ")
        cueTrig = False
        cueCount = 0
        if selectedCue == True:
            self.cueList.Select(selCue, 0)
            selectedCue = False

        message = b'<5550000000000000>' # Sets all receivers to black
        self.Send_Cue(message)
     

    def Add_New_Cue(self, cue):
        global cueContainer
        global nonitem
        xcount = self.cueList.GetItemCount() - 1
        cueContainer[1] = str(xcount + 2)
        if xcount > 0 or xcount == -1:
            self.cueList.Append(cueContainer)
        else:
            itemx = ", ".join([self.cueList.GetItem(0, col).GetText()
                               for col in range(7)])
            if itemx == nonitem:
                self.cueList.DeleteItem(xcount)
                cueContainer[1] = "1"
                self.cueList.Append(cueContainer)
            else:
                self.cueList.Append(cueContainer)
        self.cueBu3.Enable()
        self.cueBu8.Enable()

        
    def Link_Cue(self, cue):
        global cueContainer
        global cueTrig
        global cueCount
        global selCue
        global selectedCue
        if selectedCue == True and selCue > 0:
            self.cueList.SetItemBackgroundColour(selCue, (220,220,220))   
            
        if cueTrig == True:
            self.cueList.SetItem((cueCount - 1), 0, " ")
            cueCount = 0
        self.cueList.Select(selCue, 0)

        if cueTrig == True:
            cueTrig = False
            message = b'<5550000000000000>' # Sets all receivers to black
            self.Send_Cue(message)

    def Insert_Cue(self, cue):
        global cueContainer
        global cueCount
        global selCue
        global cueTrig
        xcount = self.cueList.GetItemCount() + 1
        cueContainer[1] = str(selCue)
        self.cueList.InsertItem(selCue, cueContainer[0])
        for i in range(1,8):
            self.cueList.SetItem(selCue, i, cueContainer[i])
        if cueTrig == True:
            self.cueList.SetItem((cueCount - 1), 0, " ")
            cueCount = 0
        self.cueList.Select(selCue, 0)
        for i in range(0, xcount):
            self.cueList.SetItem(i, 1, str(i + 1))
        self.cueList.Select(selCue + 1, 0)

        if cueTrig == True:
            cueTrig = False
            message = b'<5550000000000000>' # Sets all receivers to black
            self.Send_Cue(message)

    def Rep_Cue(self, cue):
        global cueContainer
        global cueCount
        global selCue
        global cueTrig
        cueContainer[1] = str(selCue + 1)
        for i in range(0, 8):
            self.cueList.SetItem(selCue, i, cueContainer[i])
        if cueTrig == True:
            self.cueList.SetItem((cueCount - 1), 0, " ")
            cueCount = 0
        self.cueList.Select(selCue, 0)

        if cueTrig == True:
            cueTrig = False
            message = b'<5550000000000000>' # Sets all receivers to black
            self.Send_Cue(message)

    def Del_Cue(self, cue):
        global cueCount
        global cueTrig
        global selCue
        global selectedCue
        xcount = self.cueList.GetItemCount() - 1
        nullo = ["", "", "", "", "", "", ""]
        if selectedCue == True:
            self.cueList.DeleteItem(selCue)
            selectedCue = False
            if xcount > 0:
                for i in range(0, xcount):
                    self.cueList.SetItem(i, 1, str(i + 1))
        else:
            if cueTrig == True:
                cueTrig = False
                self.cueList.SetItem((cueCount - 1), 0, " ")
                cueCount = 0
            if xcount > 0:
                self.cueList.DeleteItem(xcount)
            if xcount == 0:
                self.cueList.DeleteItem(xcount)
                self.cueList.Append(nullo)
                self.cueBu3.Disable()
                self.cueBu8.Disable()
        self.cueList.Update()

        message = b'<5550000000000000>' # Sets all receivers to black
        self.Send_Cue(message)

    def Rem_All_Cues(self, cue):
        global cueCount
        global cueTrig
        nullo = ["", "", "", "", "", "", ""]
        self.cueList.DeleteAllItems()
        self.cueList.Append(nullo)
        self.cueBu3.Disable()
        self.cueBu8.Disable()
        cueCount = 0
        if cueTrig == True:
            cueTrig = False
            message = b'<5550000000000000>' # Sets all receivers to black
            self.Send_Cue(message)
        cueTrig = False


    def test_Save_File(self, evt):
        with wx.FileDialog(self, "Save Cue List", wildcard="XLSX files (*.xlsx)|*.xlsx",
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fileDialog:
    
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return     # the user changed their mind
    
            # save the current contents in the file
            pathname = fileDialog.GetPath()
            
            #try:
            #    with open(pathname, 'w') as file:
            #        file.write("/ ".join(self.collist[1:]) + "\n")
            #        for row in range(self.cueList.GetItemCount()):
            #            file.write("  /".join([self.cueList.GetItem(row, col).GetText() for col in range(1, self.cueList.GetColumnCount())]) + "\n")
            #except IOError:
            #    wx.LogError("Cannot save current data in file '%s'." % pathname)

            try:
                workbook = xlsxwriter.Workbook(pathname)
                worksheet = workbook.add_worksheet()

                rowx = 0
                colx = 0

                bold = workbook.add_format({'bold': True})
                bold.set_font_size(14)
                hiLight = workbook.add_format()                
                hiLight.set_pattern(1)  # This is optional when using a solid fill.
                hiLight.set_bg_color('silver')

                headList = self.collist[1:]
                headList[4] = "Speed"

                for i in range(7):
                    worksheet.write(rowx, colx, headList[i], bold)
                    colx += 1
                
                rowx += 1
                colx = 0
                for row in range(self.cueList.GetItemCount()):
                    for col in range(self.cueList.GetColumnCount() - 1):
                        coLink = self.cueList.GetItemBackgroundColour(row)
                        if coLink == (220, 220, 220, 255):
                            worksheet.write(rowx, colx, self.cueList.GetItem(row, col + 1).GetText(), hiLight)
                        else:
                            worksheet.write(rowx, colx, self.cueList.GetItem(row, col + 1).GetText())
                        colx += 1
                    colx = 0
                    rowx += 1
                
                workbook.close()

            except IOError:
                wx.LogError("Cannot save current data in file '%s'." % pathname)



    def On_Target_Select(self, target):
        global cueContainer
        cueContainer[2] = target.GetString()
        print(cueContainer)

    def On_Mode_Select(self, mode):
        global cueContainer
        if isinstance(mode, str):
            findex = self.modelist.index(mode)
        else:
            findex = self.modelist.index(mode.GetString())
        cueContainer[3] = self.shortList[findex]
        """If Solid Color Selected."""
        if findex == 0:
            self.colorPrev1.Enable()
            self.R1Slider.Enable()
            self.G1Slider.Enable()
            self.B1Slider.Enable()
            self.colorPrev2.Disable()
            self.R2Slider.Disable()
            self.G2Slider.Disable()
            self.B2Slider.Disable()
            self.AddFX.Disable()
            self.AddSpd.Disable()
            
        """If Tap React Selected."""
        if findex == 1:
            self.colorPrev1.Disable()
            self.R1Slider.Disable()
            self.G1Slider.Disable()
            self.B1Slider.Disable()
            self.colorPrev2.Enable()
            self.R2Slider.Enable()
            self.G2Slider.Enable()
            self.B2Slider.Enable()
            self.AddFX.Disable()
            self.AddSpd.Enable()
            
        """If Effects Selected."""
        if findex == 2:
            self.colorPrev1.Disable()
            self.R1Slider.Disable()
            self.G1Slider.Disable()
            self.B1Slider.Disable()
            self.colorPrev2.Disable()
            self.R2Slider.Disable()
            self.G2Slider.Disable()
            self.B2Slider.Disable()
            self.AddFX.Enable()
            self.AddSpd.Enable()


    def Set_Effect(self, fx):
        global cueContainer
        cueContainer[4] = str(self.AddFX.GetValue())

    def Set_Speed(self, fx):
        global cueContainer
        cueContainer[5] = str(self.AddSpd.GetValue())


    def Set_Color_1(self, rgb):
        global cueContainer
        r = self.R1Slider.GetValue()
        g = self.G1Slider.GetValue()
        b = self.B1Slider.GetValue()
        self.R1Count.SetValue(str(r))
        self.G1Count.SetValue(str(g))
        self.B1Count.SetValue(str(b))
        self.colorPrev1.SetBackgroundColour((r, g, b))
        self.colorPrev1.Clear()
        cueContainer[6] = ''.join([str(r), ',', str(g), ',', str(b)])

        
    def Set_Color_2(self, rgb):
        r = self.R2Slider.GetValue()
        g = self.G2Slider.GetValue()
        b = self.B2Slider.GetValue()
        self.R2Count.SetValue(str(r))
        self.G2Count.SetValue(str(g))
        self.B2Count.SetValue(str(b))
        self.colorPrev2.SetBackgroundColour((r, g, b))
        self.colorPrev2.Clear()
        cueContainer[7] = ''.join([str(r), ',', str(g), ',', str(b)])

    def Test_Cue_Send(self, evt):
        """Function for testing currently set parameters"""
        global selCom
        header = "<555"
        footer = ">"
        group = str(self.glist.index(self.cb2.GetValue()))    # Gets Target Value
        mod = str(self.modelist.index(self.cb3.GetValue())) # Gets Mode Value
        fx = self.AddFX.GetValue() # Gets FX Value
        spd = self.AddSpd.GetValue() + 48 # Gets Speed Value
        
        color1 = ""
        color2 = ""
        sliders = [self.R1Slider, self.G1Slider, self.B1Slider, self.R2Slider, self.G2Slider, self.B2Slider]
        for i in range(6):
            a = str(int(((sliders[i].GetValue() / 100) % 10)))
            b = str(int(((sliders[i].GetValue() / 10) % 10)))
            c = str(int((sliders[i].GetValue() % 10)))
            if i < 3:
                color1 = color1 + a + b + c
            else:
                color2 = color2 + a + b + c

        """If Tap Mode Selected"""
        if mod == "1":            
            preMess = header + group + mod + fx + spd + color2 + footer
            
            
        """If Solid Color or FX Mode Selected"""
        if mod == "0" or mod == "2":
            preMess = header + group + mod + fx + spd + color1 + footer
        
        message = preMess.encode('utf-8')
        self.Send_Cue(message)


    def Update_Device(self, com):
        global gFade
        global success
        global selCom
        global modeSelect
        temp = b'<141SET'
        term = b'>'
        ser = serial.Serial(selCom, 9600, timeout=0.5, write_timeout=0.5)
        
        if modeSelect == 1:
            # Get DMX Address Value
            addnum = self.AddDMX.GetValue()
            calc = addnum - 1
            mess1 = str(int(((calc / 100) % 10))).encode('utf-8')
            mess2 = str(int(((calc / 10) % 10))).encode('utf-8')
            mess3 = str(int((calc % 10))).encode('utf-8')
            message = temp + mess1 + mess2 + mess3 + term
            try:
                ser.isOpen()
                ser.write(message)       
                bytes = ser.readline()
                test = bytes.decode("utf-8").replace('\n', '')
                if test[:7] == "1413525":
                    success = True
                ser.close()
            except SerialException:
                ser.close()
                self.DevTyp.Enable()
                self.DevTyp.SetForegroundColour((255, 255, 255))
                self.DevTyp.SetBackgroundColour((255, 0, 0))
                self.DevTyp.Clear()
                self.DevTyp.SetValue("Port Busy or Not Responding")
            
            
        if modeSelect == 2:
            # Get Group Value
            recDialog = Receiver_Dialog(None, title="Taps Receiver Detected!")
            recDialog.ShowModal()
            recDialog.Destroy()


        
        if modeSelect == 0:
            ser.close()
            self.DevTyp.Enable()
            self.DevTyp.SetForegroundColour((255, 255, 255))
            self.DevTyp.SetBackgroundColour((255, 0, 0))
            self.DevTyp.Clear()
            self.DevTyp.SetValue("Port Busy or Not Responding")

        
        
        if success == True:
            success = False
            gFade = 255
            self.greenTimer.Start(15)


    def Refresh_Dev_List(self, x):
        """Refresh the listed devices."""
        self.cb.Clear()
        ports = list(port_list.comports())
        if not ports:
            self.cb.Append("No COM Ports Available")
            self.cb.SetValue("No COM Ports Available")
            self.cb.Disable()
        else:
            self.cb.Enable()
            for p in ports:
                self.cb.Append(str(p))
                
        self.DevTyp.SetForegroundColour((0, 0, 0))
        self.DevTyp.SetBackgroundColour((255, 255, 255))
        self.DevTyp.Clear()
        self.DevTyp.Disable()
        self.AddDMX.Disable()
        self.addbu.Disable()
        self.cb2.Disable()
        self.cb3.Disable()
        self.cueBu1.Disable()
        self.cueBu7.Disable()
        self.cueBu9.Disable()
        self.colorPrev1.Disable()
        self.R1Slider.Disable()
        self.G1Slider.Disable()
        self.B1Slider.Disable()
        self.colorPrev2.Disable()
        self.R2Slider.Disable()
        self.G2Slider.Disable()
        self.B2Slider.Disable()
        self.AddFX.Disable()
        self.AddSpd.Disable()
            

    def On_Dev_Select(self, e):
        """Send message to selected COM Port to determine if valid device."""
        pname = e.GetString()
        print(pname)

        # Parse Port Number from selected port
        a = pname.split(' ')
        b = a[0]
        
        # Initialize Serial with selected port
        try:
            self.init_serial(b)
        except SerialException:
            self.DevTyp.Enable()
            self.DevTyp.SetForegroundColour((255, 255, 255))
            self.DevTyp.SetBackgroundColour((255, 0, 0))
            self.DevTyp.Clear()
            self.DevTyp.SetValue("Invalid Device")
            modeSelect = 0


    def init_serial(self, com):
        global recGroup
        global sensGroup
        global modeSelect
        global selCom
        selCom = com
        ser = serial.Serial(com, 9600, timeout=0.5, write_timeout=0.5)
        try:
            ser.isOpen()
            temp = b'<141INI>'
            ser.write(temp)        
            bytes = ser.readline()
            test = bytes.decode("utf-8").replace('\n', '')
            if test[:3] == '525':  # If Transmitter Detected
                print(test)
                self.AddDMX.Enable()
                self.addbu.Enable()
                self.DevTyp.SetForegroundColour((0, 0, 0))
                self.DevTyp.SetBackgroundColour((255, 255, 255))
                self.DevTyp.Clear()
                self.DevTyp.SetValue("Tap Shoes Transmitter")
                self.DevTyp.Enable()
                self.cb2.Enable()
                self.cb3.Enable()
                self.cueBu1.Enable()
                self.cueBu7.Enable()
                self.cueBu9.Enable()
                count = 0
                for x in range(0, 3):
                    count = count * 10 + int(test[3 + x])
                self.AddDMX.SetValue(count + 1)
                modeSelect = 1
                ser.close()
                self.On_Mode_Select(self.cb3.GetValue())            
            if test[:3] == '347':  # If receiver detected
                self.AddDMX.Disable()
                self.addbu.Enable()
                self.DevTyp.SetForegroundColour((0, 0, 0))
                self.DevTyp.SetBackgroundColour((255, 255, 255))
                self.DevTyp.Clear()
                self.DevTyp.SetValue("Tap Shoes Receiver")
                self.DevTyp.Enable()
                recGroup = str(test[3])
                sensGroup = ord(test[4]) - 48
                modeSelect = 2
                ser.close()
                recDialog = Receiver_Dialog(None, title="Taps Receiver Detected!")
                recDialog.ShowModal()
                recDialog.Destroy()
            if test[:3] not in ('525', '347'):  # If device is Invalid
                self.AddDMX.Disable()
                self.addbu.Disable()
                self.DevTyp.Enable()
                self.DevTyp.SetForegroundColour((255, 255, 255))
                self.DevTyp.SetBackgroundColour((255, 0, 0))
                self.DevTyp.Clear()
                self.DevTyp.SetValue("Invalid Device Response")
                modeSelect = 0
                ser.close()
        except SerialException:
            ser.close()
            self.DevTyp.Enable()
            self.DevTyp.SetForegroundColour((255, 255, 255))
            self.DevTyp.SetBackgroundColour((255, 0, 0))
            self.DevTyp.Clear()
            self.DevTyp.SetValue("Port Busy or Not Responding")
            modeSelect = 0

    
    def coolGreen(self, event):
        global gFade
        gFade = gFade - 5
        self.DevTyp.SetBackgroundColour((255 - gFade, 255, 255 - gFade))
        self.DevTyp.SetForegroundColour((0, gFade, 0))
        if gFade <= 0:
            gFade = 0
            self.DevTyp.SetBackgroundColour((255, 255, 255))
            self.DevTyp.SetForegroundColour((0, 0, 0))
            self.greenTimer.Stop()
        self.DevTyp.Refresh()


    def OnClose(self, event):
        global serOpen
        if serOpen == True:
            oldser = serial.Serial(selCom, 9600)
            oldser.close()       
        Quit_Program()
                    
    def OnShowPop(self, event):
        recDialog = Receiver_Dialog(None, title="Taps Receiver Detected!")
        recDialog.ShowModal()
        recDialog.Destroy()

    def saveToFile(self, event):
        saveDialog = saveFiles(None, message="Save File", wildcard=".txt", style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
        saveDialog.ShowModal()
        saveDialog.Destroy()


class Receiver_Dialog(wx.Dialog):
    """Dialog box for setting Receiver Group assignment."""
    def __init__(self, parent, **kw):
        wx.Dialog.__init__(self, parent, **kw)
        
        self.redTimer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.coolRed, self.redTimer)
        
        self.grnTimer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.coolGrn, self.grnTimer)
        
        self.rFade = 0
        self.grnFade = 0
        
        self.InitUI()
        self.SetSize(250, 233)
        self.SetTitle("Tap Receiver Detected!")
        
    def InitUI(self):
        global recGroup
        global sensGroup
        glist = ["A", "B"]
        pnl = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        sb1 = wx.StaticBox(pnl, label='Group:', size=(223, 86))
        sb2 = wx.StaticBox(pnl, label='Sensitivity:', size=(223, 54), pos=(0, 85))


        self.currLabel = wx.StaticText(pnl, label="Current Group:", pos=(20, 23))
        self.currBox = wx.TextCtrl(pnl, value=recGroup, size=(80, 22), pos=(110, 20))
        self.currBox.Disable()
        self.LabelGroup = wx.StaticText(pnl, label='Set Group:', pos=(45, 53))
        self.GSet = wx.ComboBox(pnl, value="A", size=(80, 22), pos=(110, 50), choices=glist, style=wx.CB_READONLY)

        self.sensSlide = wx.Slider(pnl, style=wx.SL_HORIZONTAL, minValue=1, maxValue=9, value=sensGroup, size=(170, 25), pos=(3, 105))
        self.sensSlide.Bind(wx.EVT_SLIDER, self.Slider_Update)
        self.sensCount = wx.TextCtrl(pnl, style=wx.TE_READONLY, value=str(sensGroup), size=(40, 22), pos=(175, 105))

        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        self.okButton = wx.Button(self, label='Set')
        closeButton = wx.Button(self, label='Close')
        hbox2.Add(self.okButton)
        hbox2.Add(closeButton, flag=wx.LEFT, border=5)

        vbox.Add(pnl, proportion=1,
            flag=wx.ALL|wx.EXPAND, border=5)
        vbox.Add(hbox2, flag=wx.ALIGN_CENTER|wx.TOP|wx.BOTTOM, border=10)

        self.SetSizer(vbox)

        self.okButton.Bind(wx.EVT_BUTTON, self.Update_Group)
        closeButton.Bind(wx.EVT_BUTTON, self.OnClose)

    def Slider_Update(self, x):
        x = self.sensSlide.GetValue()
        self.sensCount.SetValue(str(x))
        
        
    def Update_Group(self, x):
        global recGroup
        global sensGroup
        global selCom
        temp = b'<141SET'
        term = b'>'
        ser = serial.Serial(selCom, 9600, timeout=0.5, write_timeout=0.5)            
        # Get Group Value
        grouplet = self.GSet.GetValue()
        sensVal = self.sensSlide.GetValue() # Adjusts slider val to avoid sending a negative number
        sens = str(sensVal).encode("utf_8")
        mess = str(grouplet).encode("utf_8")
        message = temp + mess + sens + term
        try:
            ser.isOpen()
            ser.write(message)
            decBytes = ser.readline()
            test = decBytes.decode("utf-8").replace('\n', '')
            if test[:7] == "1413525":
                self.okButton.SetLabel("OK!")
                self.okButton.SetBackgroundColour(wx.Colour(0,255,0))
                self.currBox.Update()
                self.currBox.SetValue(grouplet)
                self.currBox.Update()
                recGroup = grouplet
                sensGroup = sensVal
            ser.close()
            self.grnTimer.Start(30)

        except SerialException:
            self.okButton.SetLabel("Failed")
            self.okButton.SetBackgroundColour(wx.Colour(255,0,0))
            self.okButton.Update()
            self.redTimer.Start(30)


    def coolRed(self, event):
        self.rFade = self.rFade + 5
        pFade = self.rFade / 255
        oneCalc = 230 * pFade
        dimRed = 255 - (25 * pFade)
        self.okButton.SetBackgroundColour(( dimRed, oneCalc, oneCalc))
        if self.rFade >= 255:
            self.rFade = 0
            self.okButton.SetLabel("Set")
            self.okButton.SetBackgroundColour(wx.NullColour)
            self.redTimer.Stop()
        self.okButton.Update()
        
    def coolGrn(self, event):
        self.grnFade = self.grnFade + 5
        pFade = self.grnFade / 255
        oneCalc = 230 * pFade
        dimGrn = 255 - (25 * pFade)
        self.okButton.SetBackgroundColour((oneCalc, dimGrn, oneCalc))
        if self.grnFade >= 255:
            self.grnFade = 0
            self.okButton.SetLabel("Set")
            self.okButton.SetBackgroundColour(wx.NullColour)
            self.grnTimer.Stop()
        self.okButton.Update()


    def OnClose(self, e):
        self.Destroy()


def Quit_Program():
    global Quit_BOOL
    Quit_BOOL = True
    sys.exit()
            
            
def main():
    ex = wx.App()
    BuildGUI(None)
    ex.MainLoop()    

if __name__ == '__main__':
    main()