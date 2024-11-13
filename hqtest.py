import sys,datetime as dt
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FC
from mpl_toolkits.basemap import Basemap
import matplotlib.pyplot as plt
import numpy as np
import win32com.client
import pythoncom
from geopy.geocoders import Nominatim
import sqlite3
import json

class T:
    def __init__(s,n,d,h,t='REQUEST',l='',dp='',r='',c='',eid='',sd=None,pri='Normal'):
        s.n,s.d,s.h,s.t,s.l,s.dp,s.r,s.c,s.eid,s.sd,s.pri=n,d,h,t,l,dp,r,c,eid,sd,pri
        s.p=0;s.st='PENDING';s.s=[]

class M(QMainWindow):
    def __init__(s):
        super().__init__()
        s.setWindowTitle('Task Manager');s.setGeometry(100,100,1200,800)
        s.task_types=['REQUEST','CONTRACT','MEETING','PROJECT']
        s.show_completed = False
        s.departments=['IT','PRODUCT MANAGEMENT','ENGINEERING','PROCUREMENT','HR','FINANCE','OPERATIONS','SALES','MARKETING']
        s.priorities=['Low','Normal','High']
        s.ts=[];s.markers=[]
        s.colors={'REQUEST':'blue','CONTRACT':'red','MEETING':'purple','PROJECT':'green'}
        s.status_colors={'PENDING':'orange','IN PROGRESS':'yellow','COMPLETED':'green','BLOCKED':'red'}
        s.init_db()  # Initialize database
        s.init_outlook()
        s.geocoder=Nominatim(user_agent="task_manager")
        s.i()
        s.load_tasks()  # Load saved tasks

    def init_db(s):
        try:
            s.conn = sqlite3.connect('taskmanager.db')
            s.cursor = s.conn.cursor()
            s.cursor.execute('''
                CREATE TABLE IF NOT EXISTS tasks (
                    id INTEGER PRIMARY KEY,
                    name TEXT, due_date TEXT, hours INTEGER,
                    task_type TEXT, location TEXT, department TEXT,
                    requestor TEXT, company TEXT, entry_id TEXT,
                    start_date TEXT, priority TEXT, progress INTEGER,
                    status TEXT, subtasks TEXT
                )
            ''')
            s.conn.commit()
        except Exception as e:
            QMessageBox.warning(s, "Database Error", f"Error initializing database: {str(e)}")

    def save_tasks(s):
        try:
            s.cursor.execute("DELETE FROM tasks")
            for t in s.ts:
                subtasks_json = json.dumps([{
                    'name': st.n, 'due_date': st.d.strftime('%Y-%m-%d %H:%M'),
                    'hours': st.h, 'task_type': st.t, 'location': st.l,
                    'department': st.dp, 'requestor': st.r, 'company': st.c,
                    'progress': st.p, 'status': st.st, 'priority': st.pri
                } for st in t.s])
                
                s.cursor.execute('''
                    INSERT INTO tasks VALUES (NULL,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ''', (t.n, t.d.strftime('%Y-%m-%d %H:%M'), t.h, t.t, t.l, t.dp,
                      t.r, t.c, t.eid, 
                      t.sd.strftime('%Y-%m-%d %H:%M') if t.sd else None,
                      t.pri, t.p, t.st, subtasks_json))
            s.conn.commit()
        except Exception as e:
            QMessageBox.warning(s, "Save Error", f"Error saving tasks: {str(e)}")

    def load_tasks(s):
        try:
            s.ts = []
            for row in s.cursor.execute("SELECT * FROM tasks"):
                task = T(
                    row[1], dt.datetime.strptime(row[2], '%Y-%m-%d %H:%M'),
                    row[3], row[4], row[5], row[6], row[7], row[8], row[9],
                    dt.datetime.strptime(row[10], '%Y-%m-%d %H:%M') if row[10] else None,
                    row[11]
                )
                task.p = row[12]
                task.st = row[13]
                if row[14]:  # subtasks
                    subtasks = json.loads(row[14])
                    for st_data in subtasks:
                        subtask = T(
                            st_data['name'],
                            dt.datetime.strptime(st_data['due_date'], '%Y-%m-%d %H:%M'),
                            st_data['hours'], st_data['task_type'],
                            st_data['location'], st_data['department'],
                            st_data['requestor'], st_data['company']
                        )
                        subtask.p = st_data['progress']
                        subtask.st = st_data['status']
                        subtask.pri = st_data['priority']
                        task.s.append(subtask)
                s.ts.append(task)
            s.ul()
        except Exception as e:
            QMessageBox.warning(s, "Load Error", f"Error loading tasks: {str(e)}")
    def init_outlook(s):
        try:
            pythoncom.CoInitialize()
            s.outlook=win32com.client.Dispatch("Outlook.Application")
            s.mapi=s.outlook.GetNamespace("MAPI")
            s.tasks_folder=s.mapi.GetDefaultFolder(13)
        except Exception as e:
            QMessageBox.warning(s,"Error","Outlook initialization failed: "+str(e))

    def ul(s):
        s.tl.clear()
        active_tasks={t:[]for t in s.task_types}
        completed_tasks=[]
        
        for t in s.ts:
            task_item=QTreeWidgetItem([t.t,t.n,
                t.sd.strftime('%Y-%m-%d %H:%M') if t.sd else '',
                t.d.strftime('%Y-%m-%d %H:%M'),
                str(t.h),t.t,t.l,t.dp,t.r,t.c,f"{t.p}%",t.st,t.pri])

            for subtask in t.s:
                subtask_item = QTreeWidgetItem([subtask.t, subtask.n,
                    subtask.sd.strftime('%Y-%m-%d %H:%M') if hasattr(subtask, 'sd') and subtask.sd else '',
                    subtask.d.strftime('%Y-%m-%d %H:%M'),
                    str(subtask.h), subtask.t, subtask.l, subtask.dp, 
                    subtask.r, subtask.c, f"{subtask.p}%", subtask.st, subtask.pri])
                task_item.addChild(subtask_item)    
            
            if t.st == 'COMPLETED':
                completed_tasks.append(task_item)
            else:
                active_tasks[t.t].append(task_item)

        # Add active tasks by category
        for cat in s.task_types:
            if active_tasks[cat]:
                cat_item=QTreeWidgetItem(s.tl,[cat])
                for task_item in active_tasks[cat]:
                    cat_item.addChild(task_item)
        
        # Add completed tasks section if toggle is on
        if s.show_completed and completed_tasks:
            completed_header=QTreeWidgetItem(s.tl,['Completed Tasks'])
            for task_item in completed_tasks:
                completed_header.addChild(task_item)
        
        s.tl.expandAll()
    def uc(s):
        [ax.clear()for ax in s.a]
        
        s.m=Basemap(projection='mill',llcrnrlat=-90,urcrnrlat=90,
                   llcrnrlon=-180,urcrnrlon=180,resolution='c',ax=s.a[0])
        s.m.drawcoastlines()
        s.m.drawcountries()
        s.m.fillcontinents(color='lightgray')
        s.m.drawmapboundary(fill_color='aqua')
        
        s.markers=[]
        for t in s.ts:
            try:
                if t.l:
                    loc=s.geocoder.geocode(t.l)
                    if loc:
                        x,y=s.m(loc.longitude,loc.latitude)
                        m,=s.a[0].plot(x,y,'o',color=s.colors.get(t.t,'blue'),
                                     markersize=10,alpha=0.7)
                        s.markers+=[m]
            except Exception as e:print(f"Geocoding error: {str(e)}")

        types=[t for t in s.task_types if any(task.t==t for task in s.ts)]
        if types:
            x=np.arange(len(types))
            completed=[sum(1 for t in s.ts if t.t==type_ and t.p==100)for type_ in types]
            in_progress=[sum(1 for t in s.ts if t.t==type_ and 0<t.p<100)for type_ in types]
            pending=[sum(1 for t in s.ts if t.t==type_ and t.p==0)for type_ in types]
            
            s.a[1].bar(x,pending,label='Pending',color='red')
            s.a[1].bar(x,in_progress,bottom=pending,label='In Progress',color='yellow')
            s.a[1].bar(x,completed,bottom=[i+j for i,j in zip(pending,in_progress)],
                      label='Completed',color='green')
            
            s.a[1].set_xticks(x)
            s.a[1].set_xticklabels(types,rotation=45)
            s.a[1].legend()
            s.a[1].set_title('Task Progress by Type')

        status_count={'PENDING':0,'IN PROGRESS':0,'COMPLETED':0,'BLOCKED':0}
        for t in s.ts:status_count[t.st]=status_count.get(t.st,0)+1
        if sum(status_count.values()):
            labels=[f"{k}\n({v})"for k,v in status_count.items()if v>0]
            sizes=[v for v in status_count.values()if v>0]
            colors=[s.status_colors[k]for k in status_count.keys()if status_count[k]>0]
            s.a[2].pie(sizes,labels=labels,colors=colors,autopct='%1.1f%%')
            s.a[2].set_title('Task Status Distribution')

        s.f.tight_layout()
        s.c.draw()

    def i(s):
        w=QWidget();s.setCentralWidget(w);l=QVBoxLayout(w)
        tb=QToolBar();s.addToolBar(tb)
        [tb.addAction(t,f)for t,f in[('Add Task',s.at),('Import Task',s.get_tasks),
                                    ('Status Report',s.sr)]]

        s.toggle_completed = QAction('Show Completed Tasks', s)
        s.toggle_completed.setCheckable(True)
        s.toggle_completed.triggered.connect(s.toggle_completed_tasks)
        tb.addAction(s.toggle_completed)                           
        s.tl=QTreeWidget()
        s.tl.setColumnCount(13)  # increased by 1
        s.tl.setHeaderLabels(['Category','Task','Start Date','Due','Hours','Type','Location',
            'Department','Requestor','Company','Progress','Status','Priority'])
        widths=[100,200,150,150,70,100,150,100,150,150,80,100,80]  # added width for Start Date
        [s.tl.setColumnWidth(i,w)for i,w in enumerate(widths)]
        s.tl.itemClicked.connect(s.st)
        l.addWidget(s.tl)
        
        s.f=plt.figure(figsize=(12,8))
        gs=s.f.add_gridspec(2,2)
        s.a=[s.f.add_subplot(gs[0,:]),
             s.f.add_subplot(gs[1,0]),
             s.f.add_subplot(gs[1,1])]
        s.c=FC(s.f)
        l.addWidget(s.c)
        
        s.m=Basemap(projection='mill',llcrnrlat=-90,urcrnrlat=90,
                   llcrnrlon=-180,urcrnrlon=180,resolution='c',ax=s.a[0])
        s.m.drawcoastlines();s.m.drawcountries();s.m.fillcontinents(color='lightgray')
        s.m.drawmapboundary(fill_color='aqua')
        s.f.canvas.draw()
        s.ul()
    
    def get_tasks(s):
        d=QDialog(s);l=QVBoxLayout(d);lw=QListWidget()
        try:
            tasks=s.tasks_folder.Items;tasks.Sort("[DueDate]",True)
            task_list=[tasks[j]for j in range(min(50,tasks.Count))]
            [lw.addItem(f"{t.Subject} - Due: {getattr(t,'DueDate','N/A')}")for t in task_list]
            if not lw.count():QMessageBox.warning(s,"No Tasks","No tasks found");return
        except Exception as e:QMessageBox.warning(s,"Error",f"Task loading error: {str(e)}");return

        l.addWidget(QLabel("Select Task:"));l.addWidget(lw)
        loc_input=QLineEdit()
        dep_combo=QComboBox();dep_combo.addItems(s.departments)
        l.addWidget(QLabel("Location:"));l.addWidget(loc_input)
        l.addWidget(QLabel("Department:"));l.addWidget(dep_combo)

        b=QDialogButtonBox(QDialogButtonBox.Ok|QDialogButtonBox.Cancel)
        b.accepted.connect(d.accept);b.rejected.connect(d.reject);l.addWidget(b)
        if d.exec_()==QDialog.Accepted and lw.currentItem():
            try:
                t=task_list[lw.currentRow()]
                requestor_email=str(getattr(t,'Mileage',''))
                location=getattr(t,'BillingInformation','') or loc_input.text()
                company=getattr(t,'Companies','').split(';')[0].strip()
                total_work=getattr(t,'TotalWork',0)
                hours=float(total_work)/60 if total_work else 1
                
                # Get actual progress and status
                percent_complete=getattr(t,'PercentComplete',0)
                if isinstance(percent_complete, float):
                    percent_complete = int(percent_complete * 100)
                percent_complete = min(100, max(0, percent_complete))
                
                # Get the actual status
                outlook_status=getattr(t,'Status',0)
                status_map = {
                    0: 'PENDING',
                    1: 'IN PROGRESS',
                    2: 'COMPLETED',
                    3: 'BLOCKED',
                    4: 'PENDING'  # deferred
                }
                task_status = status_map.get(outlook_status, 'PENDING')
                
                category=getattr(t,'Categories','REQUEST')
                if category not in s.task_types:category='REQUEST'
                importance=getattr(t,'Importance',1)
                priority_map={0:'Low',1:'Normal',2:'High'}
                priority=priority_map.get(importance,'Normal')
                
                # Get start date
                start_date=getattr(t,'StartDate',None)
                if not start_date:
                    start_date=dt.datetime.now()

                task=T(t.Subject,
                      getattr(t,'DueDate',dt.datetime.now()+dt.timedelta(days=1)),
                      hours,category,location,dep_combo.currentText(),requestor_email,
                      company,str(t.EntryID),start_date,priority)
                task.p=percent_complete
                task.st=task_status
                s.ts+=[task];s.ul();s.uc()
            except Exception as e:
                print(f"Detailed error: {str(e)}")
                QMessageBox.warning(s,"Error",f"Task creation error: {str(e)}")

    def at(s):
        d=QDialog(s);l=QVBoxLayout(d)
        i=[QLineEdit(),QSpinBox(),QCalendarWidget(),QComboBox(),QLineEdit(),
           QComboBox(),QLineEdit(),QLineEdit(),QComboBox()]
        i[1].setRange(1,100)
        i[2].setGridVisible(True)
        i[3].addItems(s.task_types)
        i[5].addItems(s.departments)
        i[8].addItems(s.priorities)
        [l.addWidget(QLabel(t))or l.addWidget(w)for w,t in zip(i,
            ['Task','Hours','Due Date','Type','Location','Department',
             'Requestor','Company','Priority'])]
        
        time_layout=QHBoxLayout()
        hour=QSpinBox();hour.setRange(0,23);hour.setValue(dt.datetime.now().hour)
        minute=QSpinBox();minute.setRange(0,59);minute.setValue(dt.datetime.now().minute)
        time_layout.addWidget(QLabel('Time:'));time_layout.addWidget(hour)
        time_layout.addWidget(QLabel(':'));time_layout.addWidget(minute)
        l.addLayout(time_layout)
        
        b=QDialogButtonBox(QDialogButtonBox.Ok|QDialogButtonBox.Cancel)
        b.accepted.connect(d.accept);b.rejected.connect(d.reject);l.addWidget(b)
        if d.exec_()==QDialog.Accepted:
            selected_date=i[2].selectedDate().toPyDate()
            due_date=dt.datetime.combine(selected_date,dt.time(hour.value(),minute.value()))
            t=T(i[0].text(),due_date,i[1].value(),i[3].currentText(),i[4].text(),
                i[5].currentText(),i[6].text(),i[7].text(),None,None,i[8].currentText())
            s.ts+=[t];s.ul();s.uc();s.save_tasks()

    def st(s,i):
        if not i.parent():return
        try:
            task_found=False
            for t in s.ts:
                if t.n==i.text(1):
                    selected_task=t
                    task_found=True
                    break
                for st in t.s:
                    if st.n==i.text(1):
                        selected_task=st
                        task_found=True
                        break
                if task_found:break
            
            if not task_found:return
            
            m=QMenu(s)
            m.addAction('Edit Task',lambda:s.edit_task(selected_task))
            m.addAction('+ Sub',lambda:s.as_(selected_task))
            m.addAction('Set %',lambda:s.sp(selected_task))
            m.addAction('Set Hours',lambda:s.sh(selected_task))
            m.addAction('Set Status',lambda:s.ss(selected_task))
            m.addAction('Set Priority',lambda:s.set_priority(selected_task))
            if hasattr(selected_task,'eid')and selected_task.eid:
                m.addAction('View Task',lambda:s.oe(selected_task))
            m.addAction('Delete Task',lambda:s.delete_task(selected_task))    
            m.exec_(QCursor.pos())
        except Exception as e:
            print(f"Menu creation error: {str(e)}")

    def edit_task(s,t):
        d=QDialog(s);l=QVBoxLayout(d)
        i=[QLineEdit(),QSpinBox(),QCalendarWidget(),QCalendarWidget(),QComboBox(),
           QLineEdit(),QComboBox(),QLineEdit(),QLineEdit(),QComboBox(),QComboBox()]
        
        # Set current values
        i[0].setText(t.n)
        i[1].setValue(int(t.h))
        i[1].setRange(1,100)
        i[2].setGridVisible(True)  # Start Date
        i[3].setGridVisible(True)  # Due Date
        i[4].addItems(s.task_types)
        i[4].setCurrentText(t.t)
        i[5].setText(t.l)
        i[6].addItems(s.departments)
        i[6].setCurrentText(t.dp)
        i[7].setText(t.r)
        i[8].setText(t.c)
        i[9].addItems(s.priorities)
        i[9].setCurrentText(t.pri)
        i[10].addItems(['PENDING','IN PROGRESS','COMPLETED','BLOCKED'])
        i[10].setCurrentText(t.st)
        
        [l.addWidget(QLabel(label))or l.addWidget(w)for w,label in zip(i,
            ['Task','Hours','Start Date','Due Date','Type','Location','Department',
             'Requestor','Company','Priority','Status'])]
        
        b=QDialogButtonBox(QDialogButtonBox.Ok|QDialogButtonBox.Cancel)
        b.accepted.connect(d.accept);b.rejected.connect(d.reject);l.addWidget(b)
        
        if d.exec_()==QDialog.Accepted:
            t.n=i[0].text()
            t.h=i[1].value()
            t.sd=i[2].selectedDate().toPyDate()
            t.d=i[3].selectedDate().toPyDate()
            t.t=i[4].currentText()
            t.l=i[5].text()
            t.dp=i[6].currentText()
            t.r=i[7].text()
            t.c=i[8].text()
            t.pri=i[9].currentText()
            t.st=i[10].currentText()
            s.ul();s.uc();s.save_tasks()
    def delete_task(s, t):
        reply = QMessageBox.question(s, 'Delete Task',
                                   f'Are you sure you want to delete task "{t.n}"?',
                                   QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            s.ts.remove(t)
            s.save_tasks()
            s.ul()
            s.uc()       
    def toggle_completed_tasks(s):
        s.show_completed = not s.show_completed
        s.toggle_completed.setText('Hide Completed Tasks' if s.show_completed else 'Show Completed Tasks')
        s.ul()        
    def sp(s,t):
        if(p:=QInputDialog.getInt(s,'Progress','%:',t.p,0,100)[0])>=0:
            t.p=p
            t.st='COMPLETED' if p==100 else 'IN PROGRESS' if p>0 else 'PENDING'
            s.ul();s.uc();s.save_tasks()

    def sh(s,t):
        if(h:=QInputDialog.getInt(s,'Hours','New Hours:',t.h,1,100)[0])>0:t.h=h;s.ul();s.uc();s.save_tasks()

    def ss(s,t):
        d=QDialog(s);l=QVBoxLayout(d)
        combo=QComboBox()
        combo.addItems(['PENDING','IN PROGRESS','COMPLETED','BLOCKED'])
        combo.setCurrentText(t.st)
        l.addWidget(QLabel("Select Status:"));l.addWidget(combo)
        b=QDialogButtonBox(QDialogButtonBox.Ok|QDialogButtonBox.Cancel)
        b.accepted.connect(d.accept);b.rejected.connect(d.reject);l.addWidget(b)
        if d.exec_()==QDialog.Accepted:t.st=combo.currentText();s.ul();s.uc();s.save_tasks()

    def set_priority(s,t):
        d=QDialog(s);l=QVBoxLayout(d)
        combo=QComboBox()
        combo.addItems(s.priorities)
        combo.setCurrentText(t.pri)
        l.addWidget(QLabel("Select Priority:"));l.addWidget(combo)
        b=QDialogButtonBox(QDialogButtonBox.Ok|QDialogButtonBox.Cancel)
        b.accepted.connect(d.accept);b.rejected.connect(d.reject);l.addWidget(b)
        if d.exec_()==QDialog.Accepted:t.pri=combo.currentText();s.ul();s.uc();s.save_tasks()

    def as_(s,t):
        d=QDialog(s);l=QVBoxLayout(d)
        i=[QLineEdit(),QSpinBox(),QCalendarWidget()]
        i[1].setRange(1,100)
        i[2].setGridVisible(True)
        [l.addWidget(QLabel(t))or l.addWidget(w)for w,t in zip(i,
            ['Task','Hours','Due Date'])]
        b=QDialogButtonBox(QDialogButtonBox.Ok|QDialogButtonBox.Cancel)
        b.accepted.connect(d.accept);b.rejected.connect(d.reject);l.addWidget(b)
        if d.exec_()==QDialog.Accepted:
            st=T(i[0].text(),i[2].selectedDate().toPyDate(),
                 i[1].value(),t.t,t.l,t.dp,t.r,t.c)
            t.s+=[st];s.ul();s.uc();s.save_tasks()

    def sr(s):
        d=QDialog(s)
        d.setWindowTitle('Status Report')
        l=QVBoxLayout(d)
        text=QTextEdit()
        report=['Status Report - '+dt.datetime.now().strftime('%Y-%m-%d %H:%M'),
                '\nTotal Tasks: '+str(len(s.ts)),
                '\nBy Status:']
        status_count={'PENDING':0,'IN PROGRESS':0,'COMPLETED':0,'BLOCKED':0}
        for t in s.ts:
            status_count[t.st]=status_count.get(t.st,0)+1
        for st,count in status_count.items():
            report.append(f'{st}: {count}')
        report.append('\nBy Type:')
        type_count={t:sum(1 for task in s.ts if task.t==t)for t in s.task_types}
        for t,count in type_count.items():
            report.append(f'{t}: {count}')
        report.append('\nUtilization:')
        total_hours=sum(t.h for t in s.ts)
        report.append(f'Total Hours: {total_hours}')
        report.append(f'Utilization: {min(100,total_hours/40*100):.1f}%')
        text.setText('\n'.join(report))
        l.addWidget(text)
        b=QDialogButtonBox(QDialogButtonBox.Ok)
        b.accepted.connect(d.accept)
        l.addWidget(b)
        d.exec_()

    def oe(s,t):
        try:
            task=s.tasks_folder.Items.Find("[EntryID]='"+t.eid+"'")
            task.Display()
        except Exception as e:
            print(f"Error opening Outlook task: {str(e)}")

if __name__=='__main__':
    app=QApplication(sys.argv)
    m=M()
    m.show()
    sys.exit(app.exec_())        