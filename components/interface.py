import PySimpleGUI as sg
from components.events import def_path
from components.events import filtering
from components.events import sheetlist
from components.events import saving_data


lista = []

def inter():
    #--------------------------------
    #---- building the interface ---- 
    #--------------------------------
 
    fr4 = [[sg.Listbox(values = [], size=(35,6), key = 'sheetslist', enable_events=True)]]
    
    fr5 = [[sg.Listbox(values = [], size=(35,4), key = 'procedurelist', enable_events=True)]]
    
    fr6 = [[sg.Frame('Test Procedure', fr4, relief="flat")],
         [sg.Frame('Test Procedure', fr5, relief="flat")]]
    
    fr1 = [[sg.Multiline("", key = "text", size = (30, 15))]]
    fr2 = [[sg.Multiline("", key = "text1", size = (30, 15))]]
    fr3 = [[sg.Multiline("", key = "text2", size = (30, 15), enable_events=True)]]

    #layout building
    layout = [[sg.Text('',key='fileaddress', size=(71,0), enable_events = True), sg.Button('OK', key = 'ok')],
                [sg.Frame("Execution Procedure", fr1, relief="flat"), sg.Frame("Expected Behavior", fr2, 
                relief="flat"), sg.Button('>', key = 'copy'), sg.Frame("Actual Behavior", fr3, relief="flat"), 
                sg.Frame('', fr6, relief="flat")],
                [sg.Input("0", key = 'current', size=(3,0)), sg.Text("/"),sg.Text("000", key = 'number'), 
                sg.Button('GO', key = 'go'),sg.Button('<<', key = 'previous'),sg.Button('>>', key = 'next')],
                [sg.Text(" "*50, key = 'lista')]]

    window = sg.Window('TESTCASE').Layout(layout) # window load
    event, values = window.Read(timeout=100) # update of the window
    #window.Element('fileaddress').Update(value = file_path) # updating the 'fileaddress' element
    file_path = 'Select a file'
    l0 = []
    l1 = {}
    sheet0 = ''

    while True:
        event, values = window.Read(timeout=100)
        window.Element('fileaddress').Update(value=file_path)
        
        if event == 'sheetslist': # lista de abas
            sheet0 = str(sheet0).replace('[','').replace(']','').replace("'",'')
            sheet = str(values['sheetslist']).replace('[','').replace(']','').replace("'",'')

            if len(l1) > 1:
                saving_data(file_path, sheet0)

            window.Element('text').Update(value='')
            window.Element('text1').Update(value='')
            window.Element('text2').Update(value='')            
            d, title = filtering(file_path = file_path, currenttest = sheet)
            window['procedurelist'](values=title)
            listbox = window['procedurelist']
            for index, i in enumerate(title):
                for j in d[i].keys():
                    if d[i][j]['sts'] == "Fail":
                        listbox.Widget.itemconfig(index, fg='red')
                        break
                    if d[i][j]['sts'] == "Pending":
                        listbox.Widget.itemconfig(index, fg='orange')
                        break
                    if d[i][j]['sts'] == "Pass":
                        listbox.Widget.itemconfig(index, fg='Green')

        
        if event == 'procedurelist': # lista de testes
            sheet0 = values['sheetslist']
            #d[proc][l1[ct]]['Actual'] = values['text2'].strip() # saving the actual behavior to dict
            window.Element('text').Update(value='')
            window.Element('text1').Update(value='')
            window.Element('text2').Update(value='')
            l0.clear()
            l1.clear()
            proc = str(values['procedurelist']).replace('[','').replace(']','').replace("'",'')
            for index, keys in enumerate(d[proc].keys()):
                l1[index + 1] = keys
                #print(l1)
            window.Element('number').Update(value=len(l1))
            window.Element('text').Update(value=d[proc][l1[1]]['step'])
            window.Element('text1').Update(value=d[proc][l1[1]]['Expected'])
            window.Element('text2').Update(value=d[proc][l1[1]]['Actual'])
            values['current'] = l1[1]
            window.Element('current').Update(value = 1)


        if event == 'ok':
            listbox = window['sheetslist']
            if file_path.endswith(".xlsx"):
                window['sheetslist'](values='')
                sheetlist1 = sheetlist(file_path)
                window['sheetslist'](values = list(sheetlist1.keys()))
                for i,k in enumerate(sheetlist1):
                    if sheetlist1[k] == "Fail":
                        listbox.Widget.itemconfig(i, fg='red')
                    if sheetlist1[k] == "Pending":
                        listbox.Widget.itemconfig(i, fg='orange')
                    if sheetlist1[k] == "Pass":
                        listbox.Widget.itemconfig(i, fg='Green')
                    

        if event == 'copy':
            window.Element('text2').Update(value=values['text1'].strip())
            d[proc][l1[ct]]['Actual'] = values['text1'].strip()
                
               
        if event == sg.WIN_CLOSED or event == 'Exit' or event == None:
            break
        

        if event == 'fileaddress':
            file_path = def_path()
            window.Element('fileaddress').Update(value=file_path)
         

        if event == 'previous':
            ct = int(values['current'])
            #d[proc][l1[ct]]['Actual'] = values['text2'].strip()
            if ct > 1:
                window.Element('text').Update(value='')
                window.Element('text1').Update(value='')
                window.Element('text2').Update(value='')
                ct = int(values['current']) - 1
                values['current'] = ct
                window.Element('current').Update(value=values['current'])
                window.Element('text').Update(value=d[proc][l1[ct]]['step'])
                window.Element('text1').Update(value=d[proc][l1[ct]]['Expected'])
                window.Element('text2').Update(value=d[proc][l1[ct]]['Actual'])


        if event == 'next':
            ct = int(values['current'])
            #d[proc][l1[ct]]['Actual'] = values['text2'].strip()
            if ct < len(l1):
                ct = int(values['current']) + 1
                window.Element('text').Update(value='')
                window.Element('text1').Update(value='')
                window.Element('text2').Update(value='')
                values['current'] = ct
                window.Element('current').Update(value=values['current'])
                window.Element('text').Update(value=d[proc][l1[ct]]['step'])
                window.Element('text1').Update(value=d[proc][l1[ct]]['Expected'])
                window.Element('text2').Update(value=d[proc][l1[ct]]['Actual'])

        
        if event == 'go':
            try:
                ct = int(values['current'])
                window.Element('text').Update(value=d['procedure'][ct]['title'])
            except:
                pass

        if event == 'text2':
            proc = str(values['procedurelist']).replace('[','').replace(']','').replace("'",'')
            ct = int(values['current'])
            d[proc][l1[ct]]['Actual'] = values['text2']

def main():
    inter()
    