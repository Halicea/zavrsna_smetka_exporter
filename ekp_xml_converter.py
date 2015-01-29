#!/usr/bin/python
# -*- coding: utf-8 -*-

import Tkinter
from Tkinter import W, E, N, S, BOTTOM,X
from tkFileDialog import askopenfilename
import tkFileDialog
import tkMessageBox
import tkSimpleDialog
import xlrd, codecs
import json
doc_types=[
    ("110",  u"Годишна сметка"),
    ("120",  u"Ревизорски извештај"),
    ("140",  u"Консолидирана годишна сметка"),
    ("150",  u"Консолидиран ревизорски извештај")]

aatypes={"110":"122", "120":"123","140":"125","150":"126"}

f_shifri = open('./shifri_na_dejnost.json', 'r').read()
shifra_na_dejnost = json.loads(f_shifri)
period = range(1,13)
period.reverse()
godina = range(2012,2021)

vid = [u"Нема промена", u"Пред промена", u"По промена"]
class simpleapp_tk(Tkinter.Tk):
    def __init__(self,parent):
        Tkinter.Tk.__init__(self,parent)
        self.parent = parent
        self.initialize()
        self.consolidated_subjects = []
        self.consButton = None
        self.consList = None

    def initialize(self):
        self.grid()
        self.dropvar = Tkinter.StringVar(self)
        self.dropvar.set(doc_types[0][1])
        self.dropvar.trace_variable('w', self.doc_type_changed)
        l = Tkinter.Label(self, text="Тип на Извештај:")
        l.grid(row=0, sticky=E)
        drop = Tkinter.OptionMenu(self,self.dropvar, None, *[v[1] for v in doc_types])
        drop.grid(column=1,row=0, sticky=E+W)

        self.periodvar = Tkinter.StringVar(self)
        self.periodvar.set(period[0])
        l = Tkinter.Label(self, text="Период:")
        l.grid(row=1, sticky=E)
        per = Tkinter.OptionMenu(self,self.periodvar, None, *period)
        per.grid(column=1,row=1, sticky=W+E)

        self.godvar = Tkinter.StringVar(self)
        self.godvar.set(godina[1])
        l = Tkinter.Label(self, text=u"Година:")
        l.grid(row=2, sticky=E)
        per = Tkinter.OptionMenu(self,self.godvar, None, *godina)
        per.grid(column=1,row=2, sticky=E+W)

        l = Tkinter.Label(self, text=u"ЕМБС:")
        l.grid(row=3, sticky=E)
        self.embs = Tkinter.Entry(self)
        self.embs.grid(column=1,row=3, columnspan=2, sticky=W+E)


        self.vid = Tkinter.StringVar(self)
        self.vid.set(vid[0])
        self.vid.trace_variable("w", self.vid_changed)
        l = Tkinter.Label(self, text=u"Статусни промени:")
        l.grid(row=4, sticky=E)
        v = Tkinter.OptionMenu(self,self.vid, None, *vid)
        v.grid(column=1,row=4, sticky=E+W)

        k = Tkinter.Label(self, text=u"Шифра на Дејност:")
        k.grid(row=7, sticky=E)
        self.dejnost = Tkinter.Entry(self)
        self.dejnost.grid(column=1,row=7, columnspan=2, sticky=W+E)
        button = Tkinter.Button(self,text=u"Отвори Годишна Сметка", command=self.OnButtonClick)
        button.grid(row=8, columnspan=3, sticky=W+E)
        self.grid_columnconfigure(0,weight=1)
        self.resizable(True,False)
        self.update()
        self.geometry(self.geometry())
    def add_consolidated(self):
      leid = tkSimpleDialog.askstring(u'ЕМБС на консолидиран субјект', u'Внеси ЕМБС на консолидиран субјект')
      if leid:
        isForeign = tkMessageBox.askyesno(u"Потврда за странско правно лице", u"Дали Консолидираниот субјект е странско правно лице?")
        lename = ""
        if isForeign:
          lename = tkSimpleDialog.askstring(u'Назив на странското правно лице', u'Венсете назив на странското правно лице')
          if not lename:
            lename = ""
        title = leid
        if(lename):
          title+="-"+lename
        self.consList.insert(0, title)
        self.consolidated_subjects.append({"LEID":leid, "LEName":lename})
    def doc_type_changed(self, name, index, mode):
      global app
      docid = None
      for k in doc_types:
        if k[1]==self.dropvar.get():
          docid= k[0]
      if docid in ["140", "150"]:
        if not self.consButton:
          self.consButton = Tkinter.Button(self, text=u"Додади консолидиран субјект", command=self.add_consolidated)
          self.consButton.grid(row=5, column=1, columnspan=1, sticky=W+E)
        if not self.consList:
          self.consList = Tkinter.Listbox(self)
          self.consList.grid(row=6, column=1)
        app.minsize(450,400)
      elif docid:
        self.consolidated_subjects = []
        if self.consButton:
          self.consButton.destroy()
        self.consButton = None
        if self.consList:
          self.consList.destroy()
        self.consList = None
        app.minsize(450,230)



    def vid_changed(self, name, index, mode):
      pass
    def OnButtonClick(self):
        self.cont = True
        conf = codecs.open('config.json', 'r', 'utf-8').read()
        conf = json.loads(conf)
        forms = conf['forms']
        base = conf['base']
        f = askopenfilename()
        if f:
            try:
                self.wb = xlrd.open_workbook(f, encoding_override="utf-8")
                final = []
                self.sheet = self.wb.sheets()[0]
                for form in forms:
                    result = []
                    self.sheet = self.wb.sheets()[form['sheet']]
                    for k in range(form['rows'][0],form['rows'][1]):
                        if k not in form['excluded_rows']:
                            row = []
                            cols = form['cols']
                            for col in cols:
                                row.append(self.sheet.row(k)[col].value)
                            if row[0]!='':
                                try:
                                    row[0] = str(int(row[0]))
                                    if row[1] or row[1]==0:
                                      row[1] = str(int(row[1]))

                                    if row[2] or row[2]==0:
                                      row[2] = str(int(row[2]))
                                    result.append(row)
                                except:
                                    pass
                            else:
                                pass
                                #print 'skipping row', row
                    if form['id']=='35':
                      if len(result)>=1:
                        if self.dejnost.get() in shifra_na_dejnost:
                          result[0][0]=shifra_na_dejnost[self.dejnost.get()]
                        else:
                          tkMessageBox.showerror(u'Невалидна шифра на дејност', u"""Внесовте непозната шифра на дејност. \nОбидете се повторно.""")
                          self.cont=False
                    if self.cont:
                      final.append({'id':form['id'], 'aops':result})
                if self.cont:
                  self.save_xml(final)
            except Exception,ex:
                self.cont = False
                tkMessageBox.showerror(u'Грешка при вчитување на документот',
                                  u"""Се случи грешка при генерирањето на XML-от.
Проверете дали го користите соодветиот ексел документ од Еуро Консалт Плус.
Доколку не можете да произведете документ, консултирајте ги Еуро Консалт Плус.""")
    def save_xml(self, forms):
        formtmpl = """      <Form ID="%s" xsi:type="form-%s">
%s
      </Form>
"""
        annualtmpl = u"""<AnnualAccount xmlns="http://e-submit.crm.com.mk/aaol" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" AATypeID="%s" DocTypeID="%s" LEID="%s" UnitID="" Year="%s" StatChangeTypeID="%s" Period="1" WorkingMonths="%s">"""

        result = u"""<?xml version="1.0" encoding="UTF-8"?>
%s
    <Operation ID="450">
%s
    </Operation>
%s
%s
    <Statement>Изјавувам, под морална, материјална и кривична одговорност, дека податоците во годишната сметка се точни и вистинити.</Statement>
</AnnualAccount>"""
        consolidatedTmpl =u"""    <ConsolidatedLEs>
      %s
    </ConsolidatedLEs>"""
        consolidatedSubjectTmpl=u"""<ConsolidatedLE %s LEID="%s"></ConsolidatedLE>"""

        template = """            <AOP ID="%s" %s xsi:type="aop-form-%s"/>"""
        auditing_company_tmpl = u"""    <AuditingCompany LEID="6061931"></AuditingCompany>
"""
        formstrs = []
        for form in forms:
          rar = []
          items = form['aops']
          for k in items:
              try:
                  current = """Current="%s" """
                  previous = """Previous="%s" """
                  values =""
                  if k[1]!="":
                    values+=current%k[1]
                  if k[2]!="":
                    values+=previous%k[2]
                  if values:
                    rar.append(template%(k[0],values, form['id']))
              except Exception, ex:
                  print ex.message
                  print k
                  raise ex
          aopsstr = '\n'.join(rar)
          formstr = formtmpl%(form['id'],form['id'],aopsstr)
          formstrs.append(formstr)
        docid = None
        for k in doc_types:
          if k[1]==self.dropvar.get():
            docid= k[0]

        period = self.periodvar.get()
        vidid = 1
        if self.vid.get():
          vidid = vid.index(self.vid.get())+1
        year = self.godvar.get()
        leid = self.embs.get()
        #DocTypeID="%s" LEID="%s" Year="%s" StatChangeTypeID="%s" Period="1" WorkingMonths="%s"
        annual_vals = annualtmpl%(aatypes[docid], docid, leid, year, vidid, period)
        consolidated_vals = ''
        auditing_company = ''
        if docid in ["140", "150"]:
          res = ''
          for k in self.consolidated_subjects:
            lename = u"LEName=\"%s\""
            if k['LEName']:
              lename=lename%k['LEName']
            else:
              lename = ''
            res+=consolidatedSubjectTmpl%(lename,k['LEID'])
          consolidated_vals = consolidatedTmpl%res
        if docid in ["120", "150"]:
          auditing_company = auditing_company_tmpl
        final = result%(annual_vals, '\n'.join(formstrs), consolidated_vals, auditing_company)
        options = {}
        options['filetypes'] = [('all files', '.*'), ('xml files', '.xml')]
        options['initialfile'] = 'zavrsna.xml'
        filename = tkFileDialog.asksaveasfilename(**options)
        if filename:
          f = codecs.open(filename, 'w', 'utf-8')
          f.write(final)
          f.close()
          app.quit()
    def OnPressEnter(self,event):
        self.OnButtonClick()
app  = None
if __name__ == "__main__":
    global app
    app = simpleapp_tk(None)
    app.minsize(450,225)
    #app.maxsize(450, 225)

    app.title(u'Еуро Консалт Плус: Конверзија во .xml')
    app.iconbitmap('favicon.ico')
    app.mainloop()
