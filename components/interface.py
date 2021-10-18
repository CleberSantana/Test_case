from queue import Empty
import PySimpleGUI as sg
from components.events import def_path
from components.events import filtering
from components.events import sheetlist
from components.events import ko, removevalue, clearlist


lista = []

def inter():
    #--------------------------------
    #---- building the interface ---- 
    #--------------------------------
 
    fr4 = [ # 4th frame, with buttons, checkbox and listbox
           [sg.Text('KO'), sg.Checkbox('', key = 'KO', enable_events = True)],
           [sg.Text('% KO:'), sg.Text('000', key = 'porcko')],
           [sg.Listbox(values = [], size=(10,5), key = 'listbox', enable_events=True)]]
    
    fr1 = [[sg.Multiline("", key = "text", size = (30, 15))]]
    fr2 = [[sg.Multiline("", key = "text1", size = (30, 15))]]
    fr3 = [[sg.Multiline("", key = "text2", size = (30, 15))]]

    #layout building
    layout = [[sg.Text('',key='fileaddress', size=(71,0), enable_events = True), sg.Button('OK', key = 'ok'), sg.Combo(values = [], size=(35,1), key = 'sheetslist', enable_events=True)],
                [sg.Frame("Execution Procedure", fr1, relief="flat"), sg.Frame("Expected Behavior", fr2, relief="flat"), sg.Button('>', key = 'copy'), sg.Frame("Actual Behavior", fr3, relief="flat"), sg.Frame('', fr4, relief="flat")],
                [sg.Input("0", key = 'current', size=(3,0)), sg.Text("/"),sg.Text("000", key = 'number'), sg.Button('GO', key = 'go'),sg.Button('<<', key = 'previous'),sg.Button('>>', key = 'next'), sg.Combo(values = [], size=(35,1), key = 'procedurelist', enable_events=True)],
                [sg.Text(" "*50, key = 'lista')]]

    window = sg.Window('TESTCASE').Layout(layout) # window load
    event, values = window.Read(timeout=100) # update of the window
    #window.Element('fileaddress').Update(value = file_path) # updating the 'fileaddress' element
    file_path = 'Select a file'
    l0 = []
    l1 = {}
    while True:
        event, values = window.Read(timeout=100)
        window.Element('fileaddress').Update(value=file_path)
        
        if event == 'sheetslist':
            d, title = filtering(file_path = file_path, currenttest = values['sheetslist'])
            window['procedurelist'](values=title)
            #print(d)
        
        if event == 'procedurelist':
            window.Element('text').Update(value='')
            window.Element('text1').Update(value='')
            #print(d[values['procedurelist']])
            #window.Element('number').Update(value=len(d.keys()))
            l0.clear()
            l1.clear()
            for index, keys in enumerate(d[values['procedurelist']].keys()):
                l1[index + 1] = keys
                #print(l1)
            window.Element('number').Update(value=len(l1))
            window.Element('text').Update(value=d[values['procedurelist']][l1[1]]['step'])
            window.Element('text1').Update(value=d[values['procedurelist']][l1[1]]['Expected'])
            window.Element('text2').Update(value=d[values['procedurelist']][l1[1]]['Actual'])
            values['current'] = l1[1]
            window.Element('current').Update(value = 1)
            window.Element('porcko').Update(value = '0')
            window.Element('KO').Update(value = False)
            window['listbox'](values = '')
            '''if len(list1) > 0:
                list = ko ('ct','1')
                list = []
                list.extend(list1)
                window['listbox'](values=list)
            else:'''


        if event == 'ok':
            if file_path.endswith(".xlsx"):
                window['listbox'](values='')
                sheetlist1 = sheetlist(file_path)
                #print(sheetlist1)
                window['sheetslist'](values=sheetlist1)
                #window.Element('sheetslist').Update(value=sheetlist1)
                '''window.Element('number').Update(value=len(d.keys()))
                window.Element('text').Update(value=d['1']['title'])
                window.Element('text1').Update(value=d['1']['steps'])
                values['current'] = 1
                window.Element('current').Update(value=1)
                window.Element('porcko').Update(value='0')
                window.Element('KO').Update(value=False)
                window['listbox'](values='')
                if len(list1) > 0:
                    list = ko ('ct','1')
                    list = []
                    list.extend(list1)
                    window['listbox'](values=list)
                else:
                    list = ko ('ct','1')'''
                
                

            #sg.Popup('The file {} has been filtered'.format(file_path))
            
               
        if event == sg.WIN_CLOSED or event == 'Exit' or event == None:
            break
        
        if event == 'fileaddress':
            file_path = def_path()
            window.Element('fileaddress').Update(value=file_path)
         
        if event == 'previous':
            ct = int(values['current'])
            if ct > 1:
                window.Element('text').Update(value='')
                window.Element('text1').Update(value='')
                window.Element('text2').Update(value='')
                ct = int(values['current']) - 1
                values['current'] = ct
                window.Element('current').Update(value=values['current'])
                window.Element('text').Update(value=d[values['procedurelist']][l1[ct]]['step'])
                window.Element('text1').Update(value=d[values['procedurelist']][l1[ct]]['Expected'])
                window.Element('text2').Update(value=d[values['procedurelist']][l1[ct]]['Actual'])
                '''if d[ct-1]['sts'] == 'KO':
                    window.Element('KO').Update(value=True)
                else:
                    window.Element('KO').Update(value=False)'''

        if event == 'next':
            ct = int(values['current'])
            if ct < len(l1):
                ct = int(values['current']) + 1
                window.Element('text').Update(value='')
                window.Element('text1').Update(value='')
                window.Element('text2').Update(value='')
                values['current'] = ct
                window.Element('current').Update(value=values['current'])
                window.Element('text').Update(value=d[values['procedurelist']][l1[ct]]['step'])
                window.Element('text1').Update(value=d[values['procedurelist']][l1[ct]]['Expected'])
                window.Element('text2').Update(value=d[values['procedurelist']][l1[ct]]['Actual'])
                '''if d[ct]['sts'] == 'KO':
                    window.Element('KO').Update(value=True)
                elif d[ct]['sts'] == 'OK':
                    window.Element('KO').Update(value=False)'''
        
        if event == 'go':
            try:
                ct = int(values['current'])
                window.Element('text').Update(value=d['procedure'][ct]['title'])
                '''if d[ct-1]['sts'] == 'KO':
                    window.Element('KO').Update(value=True)
                elif d[ct-1]['sts'] == 'OK':
                    window.Element('KO').Update(value=False)'''
            except:
                pass

        if event == 'KO':
            ct = int(values['current']) # assign the current value to variable ct
            try:
                if d[str(ct)]['sts'] == 'OK': # comparing the value of the dict
                    d[str(ct)]['sts'] = 'KO' # changing the value of the dict for the specific spec
                    window.Element('KO').Update(value=True) # updating the element 'KO' (checkbox)
                    list = ko (ct, '0') # call the function 'list' to return the updated list of KOs
                    window.Element('porcko').Update(value=round(100*(len(list)/len(d.keys())))) # update the element 'porcko' with the percentage of the KOs 
                elif d[str(ct)]['sts'] == 'KO':
                    d[str(ct)]['sts'] = 'OK' 
                    window.Element('KO').Update(value=False)
                    list = removevalue(ct)
                    window.Element('porcko').Update(value=round(100*(len(list)/len(d.keys()))))
                #print(list)
                window['listbox'](values=list)
            except:
                pass
        
        if event == 'listbox':
            try:
                ct = int(str(values['listbox']).replace('[','').replace(']',''))
                values['current'] = ct
                window.Element('current').Update(value=values['current'])
                window.Element('text').Update(value=d[str(ct)]['title'])
                if d[str(ct)]['sts'] == 'KO':
                    window.Element('KO').Update(value=True)
                elif d[str(ct)]['sts'] == 'OK':
                    window.Element('KO').Update(value=False)
            except:
                pass
            
def main():
    inter()