from numpy import unique
import pandas as pd
from tkinter import Tk, Label, Entry, Button,filedialog ,StringVar
from matplotlib.pyplot import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os

def Fcfs(eo):
    maxval= []
    for i in range(0, len(e)):
        maxval.append(int(e[i].get())) # adding the element

    global df2
    df2 = pd.DataFrame([maxval],columns=Oplist)
    IDlist=df['ID'].unique()
    for id in IDlist:
        op=df['Option1'].where(df['ID']==id).dropna()
        if df2[op].values[0]<=0 :
              op=df['Option2'].where(df['ID']==id).dropna()
              if df2[op].values[0]<=0:
                  op=df['Option3'].where(df['ID']==id).dropna()
                  if df2[op].values[0]<=0:
                      op='None'
        df.loc[df['ID'] == id, 'SelectedOption'] = op
        df2[op]-=1

    df.loc[df['Option1']==df['SelectedOption'],'Pref']=1
    df.loc[df['Option2']==df['SelectedOption'],'Pref']=2
    df.loc[df['Option3']==df['SelectedOption'],'Pref']=3
    df.loc[df['SelectedOption']=='None','Pref']=0
    df.to_excel('resultsbyid.xlsx',columns=['ID','Name','SelectedOption','Pref'] ,index=False)

    dfs=df.groupby(['SelectedOption'])
    dfp=df.groupby(['Pref']).size()

    diffoption = df['SelectedOption'].unique()
    while(1):
        if(eo==1) :
            # Split diff sheets
            df1=[]
            i=0
            with pd.ExcelWriter('resultsbyoption.xlsx') as writer: # pylint: disable=abstract-class-instantiated
             for value in diffoption:
                df1.append(dfs.get_group(value))
                df1[i].to_excel(writer,sheet_name=value[:30],index=False,columns=['ID','Name'])
                i+=1
            break
        elif(eo==2) :
            # Split diff excel
            for value in diffoption:
                df1 = dfs.get_group(value)
                output_file_name = str(value) + ".xlsx"
                df1.to_excel(output_file_name,columns=['ID','Name'],index=False)
            break

    df2=df2.T
    df2.columns=['Remaining']
    df2['Max']=maxval
    df2['Count']=df2['Max']-df2['Remaining']
    del df2['Max']
    df2 = df2[['Count','Remaining']]

    figure1 = Figure(figsize=(17,4), dpi=90)
    ax1 = figure1.add_subplot(111)
    bar1 = FigureCanvasTkAgg(figure1, root)
    bar1.get_tk_widget().grid(row=len(Oplist)+1,column=0,columnspan=40)
    #plot count
    p1=df2.plot(kind='barh',stacked='True',ax=ax1)

    figure2 = Figure(figsize=(17,2.5), dpi=90)
    ax2 = figure2.add_subplot(111)
    bar2 = FigureCanvasTkAgg(figure2, root)
    bar2.get_tk_widget().grid(row=len(Oplist)+2,column=0,columnspan=40)
    p2=dfp.plot(kind='barh',ax=ax2)

    for p in p1.patches:
        p1.annotate(str(p.get_width()), (p.get_x()+p.get_width()/2,p.get_y() +p.get_height()/4))

    for p in p2.patches:
        p2.annotate(str(p.get_width()), (p.get_width()/2, p.get_y()+p.get_height()/4))

def open_file():
    global file
    file = filedialog.askopenfilename()
    w = Label(root, text=file,bg='white')
    os.chdir(os.path.dirname(file))
    w.grid(row=0, column=0)
    global df
    df= pd.read_excel(file)
    global Oplist
    Oplist=unique(df[['Option1', 'Option2','Option3']].values).tolist()
    Oplist.append('None')
    global e
    e=[]
        # iterating till the range
    for i in range(0, len(Oplist)):
            Label(root, text=Oplist[i],bg='white').grid(row=i+1)
            e.append(Entry(root,bg='white',textvariable=StringVar(root, value='20')))
            e[i].grid(row=i+1,column=1,columnspan=3)
    e[-1].delete(-1)
    e[-1].insert(-1,2)
global root
root = Tk()
root.state('zoomed')
root.configure(background='white')
root.title("FCFS by Ahmed Hussein")
openbtn = Button(root, text ='Open', command =open_file,bg='white')
openbtn.grid(row=0, column=1)
fcfs1btn = Button(root, text ='FCFS Diff sheet', command = lambda: Fcfs(1),bg='white' )
fcfs1btn.grid(row=0, column=2)
fcfs2btn = Button(root, text ='FCFS Diff Excel', command = lambda: Fcfs(2) ,bg='white' )
fcfs2btn.grid(row=0, column=3)

root.mainloop()
