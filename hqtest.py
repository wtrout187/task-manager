import sys,datetime as dt
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FC
import matplotlib.pyplot as plt
import numpy as np
import win32com.client
import pythoncom
from geopy.geocoders import Nominatim
import sqlite3
import json
from mpl_toolkits.mplot3d import Axes3D
import cartopy.crs as ccrs
from matplotlib.animation import FuncAnimation
import numpy as np
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.image import imread
import os
import cartopy.feature as cfeature
import numpy as np
from PyQt5.QtCore import QTimer
import networkx as nx, shapely.wkt

class T:
    def __init__(s,n,d,h,t='REQUEST',l='',dp='',r='',c='',eid='',sd=None,pri='Normal'):
        s.n,s.d,s.h,s.t,s.l,s.dp,s.r,s.c,s.eid,s.sd,s.pri=n,d,h,t,l,dp,r,c,eid,sd,pri
        s.p=0;s.st='PENDING';s.s=[]

class M(QMainWindow):
    def __init__(s):
        super().__init__()
        s.setWindowTitle('Task Manager')
        s.setGeometry(100,100,1200,800)
        s.geocode_cache = {}
        s.locations_changed = False
        s.pomodoro_timer = QTimer()
        s.pomodoro_time = 25 * 60  # 25 minutes in seconds
        s.pomodoro_active = False
        s.button_layout = QHBoxLayout()
        s.weekly_button = QPushButton("Weekly")
        s.monthly_button = QPushButton("Monthly")
        s.task_types=['REQUEST','CONTRACT','MEETING','PROJECT']
        s.show_completed = False
        s.departments=['IT','PRODUCT MANAGEMENT','ENGINEERING','PROCUREMENT','HR','FINANCE','OPERATIONS','SALES','MARKETING']
        s.priorities=['Low','Normal','High']
        s.ts=[];s.markers=[]
        s.colors={'REQUEST':'blue','CONTRACT':'red','MEETING':'purple','PROJECT':'green'}
        s.status_colors={'PENDING':'orange','IN PROGRESS':'yellow','COMPLETED':'green','BLOCKED':'red'}
        s.init_db()  # Initialize database
        s.init_outlook()
        s.geocoder=Nominatim(user_agent="task_manager", timeout=10)
        
        # Create tree widget before calling i()
        s.tl = QTreeWidget()
        s.tl.setColumnCount(12)
        s.tl.setHeaderLabels(['Category','Task','Start Date','Due','Hours', 'Type', 'Location',
            'Department','Requestor','Company','Progress','Status','Priority'])
        widths=[100,200,150,150,70,100,150,100,150,150,80,100,80]
        [s.tl.setColumnWidth(i,w) for i,w in enumerate(widths)]
        s.tl.itemClicked.connect(s.st)
        
        s.i()
        s.load_tasks()  # Load saved tasks saved tasks
    def init_pomodoro(s):
        timer_widget = QWidget()
        timer_layout = QHBoxLayout(timer_widget)
        s.timer_label = QLabel("25:00")
        s.timer_label.setStyleSheet("color: #00ffff;")
        start_button = QPushButton("Start")
        start_button.clicked.connect(s.toggle_pomodoro)
        timer_layout.addWidget(s.timer_label)
        timer_layout.addWidget(start_button)
        s.pomodoro_timer.timeout.connect(s.update_pomodoro)
        return timer_widget

    def toggle_pomodoro(s):
        s.pomodoro_timer.timeout.connect(s.update_pomodoro)
        if s.pomodoro_active:
            s.pomodoro_timer.stop()
            s.pomodoro_active = False
        else:
            s.pomodoro_timer.start(1000)  # Update every second
            s.pomodoro_active = True

    def update_pomodoro(s):
        s.pomodoro_time -= 1
        minutes = s.pomodoro_time // 60
        seconds = s.pomodoro_time % 60
        s.timer_label.setText(f"{minutes:02d}:{seconds:02d}")
        if s.pomodoro_time <= 0:
            s.pomodoro_timer.stop()
            s.pomodoro_time = 25 * 60
            QMessageBox.information(s, "Pomodoro", "Time's up!")

        #s.c.draw()
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
        s.tl.setUpdatesEnabled(False)    
        s.tl.clear()    
        s.tl.setColumnCount(13)
        s.tl.clear()
        active_tasks={t:[]for t in s.task_types}
        completed_tasks=[]
        
        
        for t in s.ts:
            task_widget = QWidget()
            task_layout = QVBoxLayout(task_widget)
            task_layout.setAlignment(Qt.AlignCenter)
            progress_bar = QProgressBar()
            progress_bar.setValue(t.p)
            progress_bar.setStyleSheet("""
                QProgressBar {
                    border: none;
                    background-color: #004d4d;
                    height: 4px;
                }
                QProgressBar::chunk {
                    background-color: #00ffff;
                }
            """)
            task_layout.addWidget(progress_bar)
            task_item=QTreeWidgetItem([t.t,t.n,
                t.sd.strftime('%Y-%m-%d %H:%M') if t.sd else '',
                t.d.strftime('%Y-%m-%d %H:%M'),
                str(t.h),t.t,t.l,t.dp,t.r,t.c,f"{t.p}%",t.st,t.pri])
            progress_widget = QProgressBar()
            progress_widget.setValue(t.p)
            progress_widget.setTextVisible(True)
            progress_widget.setStyleSheet("""
                    QProgressBar {
                        border: 1px solid #00ffff;
                        border-radius: 2px;
                        text-align: center;
                        background-color: #001a1a;
                    }
                    QProgressBar::chunk {
                        background-color: #00ffff;
                    }
                """)

            s.tl.setItemWidget(task_item, 10, progress_widget)
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
        s.tl.setUpdatesEnabled(True)
    def uc(s):
        # Add color-coded markers for different partner types
        partner_colors = {
            'Distributor': 'yellow',
            'Customer': 'green', 
            'Reseller': 'purple',
            'Vendor': 'orange',
            'Facility': 'red'
        }
        
        

         # Clear previous plots
        [ax.clear() for ax in s.a]
    
        
        # Remove box and axes
        s.a[0].set_axis_off()
        s.a[0].set_proj_type('ortho')  # Use orthographic projection
        
        
        
        # Create latitude and longitude lines
        phi = np.linspace(0, 2*np.pi, 50)
        theta = np.linspace(-np.pi/2, np.pi/2, 50)
        
        # Create latitude circles
        for lat in np.linspace(-np.pi/2, np.pi/2, 8):
            x = np.cos(lat) * np.cos(phi)
            y = np.cos(lat) * np.sin(phi)
            z = np.sin(lat) * np.ones_like(phi)
            s.a[0].plot(x, y, z, color='cyan', alpha=0.2, linewidth=0.5)  # alpha and linewidth

        # Create longitude lines
        for lon in np.linspace(0, 2*np.pi, 16):
            x = np.cos(theta) * np.cos(lon)
            y = np.cos(theta) * np.sin(lon)
            z = np.sin(theta)
            s.a[0].plot(x, y, z, color='cyan', alpha=0.2, linewidth=0.5)  #alpha and linewidth
        
        # Plot continents
        land = cfeature.NaturalEarthFeature("physical", "land", "110m")
        for geom in land.geometries():
            try:
                if hasattr(geom, 'exterior'):
                    # Handle polygon geometries
                    coords = np.array(geom.exterior.coords)
                    if len(coords) > 0:
                        lats = coords[:, 1]
                        lons = coords[:, 0]
                        x = np.cos(np.radians(lats)) * np.cos(np.radians(lons))
                        y = np.cos(np.radians(lats)) * np.sin(np.radians(lons))
                        z = np.sin(np.radians(lats))
                        s.a[0].plot(x, y, z, color='cyan', alpha=0.3, linewidth=0.8)
                        
                    # Handle interior rings (holes)
                    for interior in geom.interiors:
                        coords = np.array(interior.coords)
                        lats = coords[:, 1]
                        lons = coords[:, 0]
                        x = np.cos(np.radians(lats)) * np.cos(np.radians(lons))
                        y = np.cos(np.radians(lats)) * np.sin(np.radians(lons))
                        z = np.sin(np.radians(lats))
                        s.a[0].plot(x, y, z, color='cyan', alpha=0.3, linewidth=0.8)
                
                elif hasattr(geom, 'coords'):
                    # Handle line geometries
                    coords = np.array(geom.coords)
                    if len(coords) > 0:
                        lats = coords[:, 1]
                        lons = coords[:, 0]
                        x = np.cos(np.radians(lats)) * np.cos(np.radians(lons))
                        y = np.cos(np.radians(lats)) * np.sin(np.radians(lons))
                        z = np.sin(np.radians(lats))
                        s.a[0].plot(x, y, z, color='cyan', alpha=0.8, linewidth=1.5) 
            except Exception as e:
                continue  # Skip problematic geometries
            if s.locations_changed:
        # Globe visualization code
                s.locations_changed = False
            else:
                # Skip globe update if no locations changed
                pass    
        # Set headquarters location (Louisville, CO)
        hq_lat, hq_lon = 39.9778, -105.1319
        hq_x = np.cos(np.radians(hq_lat)) * np.cos(np.radians(hq_lon))
        hq_y = np.cos(np.radians(hq_lat)) * np.sin(np.radians(hq_lon))
        hq_z = np.sin(np.radians(hq_lat))
        
        # Plot HQ with smaller marker
        def pulse(frame):
            base_size = 15  # Base marker size
            pulse_size = base_size + 5 * np.sin(frame/10)  # Pulsing effect
            glow = s.a[0].scatter([hq_x], [hq_y], [hq_z], 
                                color='cyan', s=pulse_size*2,  # Larger glow
                                alpha=0.2, marker='o')
            marker = s.a[0].scatter([hq_x], [hq_y], [hq_z], 
                                color='white', s=pulse_size,  # Core marker
                                alpha=1, marker='o', 
                                edgecolor='cyan')
            return glow, marker
        
        # Plot task locations and connections
        for t in s.ts:
            if t.l:
                try:
                    if t.l in s.geocode_cache:
                        loc = s.geocode_cache[t.l]
                    else:
                        loc = s.geocoder.geocode(t.l)
                        s.geocode_cache[t.l] = loc  # Cache the result
                    loc = s.geocoder.geocode(t.l)
                    partner_type = t.c  # Or however you store partner type
                    color = partner_colors.get(partner_type, 'cyan')
                    if loc:
                        lat, lon = loc.latitude, loc.longitude
                        x = np.cos(np.radians(lat)) * np.cos(np.radians(lon))
                        y = np.cos(np.radians(lat)) * np.sin(np.radians(lon))
                        z = np.sin(np.radians(lat))
                        
                        # Plot location with smaller marker
                        s.a[0].scatter([x], [y], [z], 
                                    color='white', s=10, 
                                    alpha=0.8, marker='o',
                                    edgecolor=s.colors.get(t.t, 'cyan'))
                        
                        # Draw single connection line with gradient
                        line_points = 25
                        alphas = np.linspace(0.05, 0.3, line_points)  # Slightly more subtle gradient
                        x_coords = np.linspace(hq_x, x, line_points)
                        y_coords = np.linspace(hq_y, y, line_points)
                        z_coords = np.linspace(hq_z, z, line_points)

                        s.a[0].plot(x_coords, y_coords, z_coords,
                                    color='cyan', alpha=0.8, linewidth=0.8)  # Slightly thinner lines
                
                except Exception as e:
                    print(f"Geocoding error: {str(e)}")
        
       # Set view and add rotation animation
        s.a[0].view_init(elev=20, azim=45)
        s.a[0].dist = 3  # Adjust camera distance
        globe_legend_elements = [
            plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='cyan', 
                    label='Headquarters', markersize=8),
            plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='white',
                    label='Partner Location', markersize=8),
            plt.Line2D([0], [0], color='cyan', label='Connection', alpha=0.3)
        ]
        # Recreate the legend after clearing the axes
        s.a[0].legend(
            handles=globe_legend_elements,
            loc='center left',  # Adjust location as needed
            bbox_to_anchor=(1.1, 0.5),  # Adjust placement relative to the plot
            ncol=1,  # Ensure vertical layout
            fontsize=8,
            frameon=False
        )
        
        def rotate_and_pulse(frame):
            s.a[0].view_init(elev=25, azim=frame)
            pulse(frame)
            return s.a[0],

        # Increase frames for smoother rotation
        s.anim = FuncAnimation(s.f, rotate_and_pulse,
                                frames=100,  # globe rotation smoothing
                                interval=500,  # Lower interval = faster animation
                                blit=True,
                                cache_frame_data=True,
                                repeat=True)
        s.f.set_dpi(100)  # Lower DPI
        plt.rcParams['figure.max_open_warning'] = 50
        # Clear and recreate bottom chart
        # Bottom chart - Combined Utilization and Task Distribution
        s.a[1].clear()  # Clear the old utilization plot
        #s.a[2].clear()  # placeholder for additional charts/plots clearing
               
        
    def i(s):
        # Create main widget and layout
        w = QWidget()
        s.setCentralWidget(w)
        main_layout = QVBoxLayout(w)

        
        # Create figure and subplots first
        s.f = plt.figure(figsize=(16, 12))
        s.gs = s.f.add_gridspec(2, 1, height_ratios=[1.4, 1]) 
        
        s.a = [
            s.f.add_subplot(s.gs[0], projection='3d'),    # Globe in top left
            s.f.add_subplot(s.gs[1], projection='3d'),    # Asset utilization Top Right
            #s.f.add_subplot(s.gs[1, 1])                      # Task distribution bottom right
        ]
        s.c = FC(s.f)  # Create canvas after figure and subplots
         # Adjust subplot positions
        s.a[0].set_position([0.1, 0.55, 0.8, 0.60])   # Globe
        s.a[1].set_position([0.1, 0.05, 0.8, 0.4])   # Asset utilization
        #s.a[2].set_position([0.55, 0.1, 0.4, 0.35]) 
        s.uc()
        # Add a legend
        globe_legend_elements = [
            plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='cyan', 
                    label='Headquarters', markersize=8),
            plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='white',
                    label='Partner Location', markersize=8),
            plt.Line2D([0], [0], color='cyan', label='Connection', alpha=0.3)
        ]
        legend = s.a[0].legend(handles=globe_legend_elements,
                      loc='center left',
                      bbox_to_anchor=(1.15, 0.5),
                      ncol=1,
                      fontsize=12,
                      frameon=False)
        
        s.weekly_button = QPushButton("Weekly")
        s.monthly_button = QPushButton("Monthly")

        def update_plot(timeframe):
    
            s.a[1].cla()  # Clear the axes

            now = dt.datetime.now()
            if timeframe == "weekly":
                start_date = now - dt.timedelta(days=now.weekday())
                end_date = start_date + dt.timedelta(days=6)
                available_hours = 40
                date_range = f"Week of {start_date.strftime('%m/%d/%Y')}-{end_date.strftime('%m/%d/%Y')}"
            elif timeframe == "monthly":
                start_date = now.replace(day=1)
                end_date = (now + dt.timedelta(days=32)).replace(day=1) - dt.timedelta(days=1)
                available_hours = 40 * (end_date.day / 7)
                date_range = f"Month of {start_date.strftime('%m/%d/%Y')}-{end_date.strftime('%m/%d/%Y')}"

            days = (end_date - start_date).days + 1
            tasks_in_period = [t for t in s.ts if start_date <= t.d <= end_date]

            # Create smooth grid for contour plot
            x = np.linspace(0, days-1, 100)
            y = np.linspace(0, 100, 100)
            X, Y = np.meshgrid(x, y)
            Z = np.zeros_like(X)

            # Calculate task statistics and create distribution
            task_stats = {task_type: {'count': 0, 'hours': 0} for task_type in s.task_types}
            Z_by_type = {task_type: np.zeros_like(X) for task_type in s.task_types}
            
            # Create distributions for each task type
            for task in tasks_in_period:
                task_day = (task.d - start_date).days
                task_hours = task.h + sum(st.h for st in task.s)
                task_stats[task.t]['count'] += 1
                task_stats[task.t]['hours'] += task_hours
                
                # Create distribution based on task hours
                amplitude = (task_hours / available_hours) * 100
                sigma_x = days/6
                sigma_y = 15
                Z_by_type[task.t] += amplitude * np.exp(-((X - task_day)**2/(2*sigma_x**2) + (Y - 50)**2/(2*sigma_y**2)))

            # Calculate total hours and utilization
            total_hours = sum(stats['hours'] for stats in task_stats.values())
            utilization = min(100, (total_hours / available_hours) * 100) if available_hours > 0 else 0

            # Create combined surface with proper color distribution
            Z_total = np.zeros_like(X)
            active_types = [t for t in s.task_types if task_stats[t]['hours'] > 0]
            
            if active_types:
                # Create color gradient based on task distribution
                colors = []
                for t in active_types:
                    weight = task_stats[t]['hours'] / total_hours if total_hours > 0 else 0
                    colors.extend([s.colors[t]] * int(weight * 256))
                
                # Ensure we have at least 2 colors
                while len(colors) < 2:
                    colors.append(colors[0] if colors else 'blue')
                
                custom_cmap = LinearSegmentedColormap.from_list('custom', colors, N=256)
                
                # Combine all task distributions
                for task_type in active_types:
                    Z_total += Z_by_type[task_type]
                
                # Plot single surface with color gradient
                surf = s.a[1].plot_surface(X, Y, Z_total,
                                            cmap=custom_cmap,
                                            alpha=0.8,
                                            antialiased=True,  # Smooth edges
                                            shade=True,        # Add shading
                                            rcount=100,        # Increase resolution
                                            ccount=100)        # Increase resolution

            # Add legend with detailed task information
            legend_elements = []
            for task_type in s.task_types:
                if task_stats[task_type]['hours'] > 0:
                    percentage = (task_stats[task_type]['hours'] / total_hours * 100) if total_hours > 0 else 0
                    label = f"{task_type}: {task_stats[task_type]['count']} tasks, {task_stats[task_type]['hours']:.1f}h ({percentage:.1f}%)"
                    legend_elements.append((s.colors[task_type], label))

            # Position legend outside the plot
            legend = s.a[1].legend(
                [plt.Line2D([1], [0], color=color, lw=2) for color, _ in legend_elements],
                [label for _, label in legend_elements],
                loc='center left',  # Position to the left
                bbox_to_anchor=(1.1, 0.5),  # Adjust outside the plot area
                ncol=1,  # Ensure vertical layout
                fontsize=8,  # Match title font size
                frameon=False  # Remove frame for cleaner look
)        

            # Configure the plot
            s.a[1].set_xlabel('Days')
            s.a[1].set_ylabel('Utilization %')
            s.a[1].set_zlabel('Task Distribution')
            s.a[1].set_title(f'Asset Utilization ({date_range}): {utilization:.1f}%',
                            fontsize=8,
                            pad=20)

            # Remove background elements
            s.a[1].xaxis.pane.fill = False
            s.a[1].yaxis.pane.fill = False
            s.a[1].zaxis.pane.fill = False
            s.a[1].xaxis.pane.set_edgecolor('none')
            s.a[1].yaxis.pane.set_edgecolor('none')
            s.a[1].zaxis.pane.set_edgecolor('none')
            s.a[1].grid(False)

            # Set view angle and limits
            s.a[1].view_init(elev=30, azim=45)
            s.a[1].set_xlim(0, days-1)
            s.a[1].set_ylim(0, 100)

            # Set z-limit with proper scaling
            max_z = np.max(Z_total) if np.any(Z_total) else 100
            if max_z > 0:
                s.a[1].set_zlim(0, max_z * 1.1)
            else:
                s.a[1].set_zlim(0, 100)

            s.c.draw()

        # Create task container with centered layout
        task_container = QWidget()
        task_layout = QVBoxLayout(task_container)
        task_layout.addWidget(s.tl)
        task_layout.setAlignment(Qt.AlignCenter)
        task_layout.setContentsMargins(10, 10, 10, 10)
        
        # Add widgets to main layout
        main_layout.addWidget(s.c)  # Canvas with graphs at top
        
        # Integrate button_layout into your main layout
        button_widget = QWidget()
        button_widget.setLayout(s.button_layout)  # button_layout from uc()
        main_layout.addWidget(button_widget)  # Add before main_splitter
        # Create splitter for better layout control
        splitter = QSplitter(Qt.Vertical)
        splitter.addWidget(s.c)
        splitter.addWidget(task_container)
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 1)
        # Set task list size and position
        task_container.setMinimumHeight(400)
        task_container.setMaximumHeight(600)
        # Create left panel for task list
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.addWidget(s.tl)

        # Create right panel for graphs
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.addWidget(s.c)

        # Create horizontal splitter
        h_splitter = QSplitter(Qt.Horizontal)
        h_splitter.addWidget(left_panel)
        h_splitter.addWidget(right_panel)
        h_splitter.setStretchFactor(0, 1)  # Task list gets 1 part
        h_splitter.setStretchFactor(1, 2)  # Graphs get 2 parts

        # Create main horizontal splitter
        main_splitter = QSplitter(Qt.Horizontal)

        # Left panel for task list
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.addWidget(s.tl)
        left_panel.setLayout(left_layout)

        # Right panel for graphs
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.addWidget(s.c)
        right_panel.setLayout(right_layout)

        # Add panels to splitter
        main_splitter.addWidget(left_panel)
        main_splitter.addWidget(right_panel)
        main_splitter.setStretchFactor(0, 1)  # Task list
        main_splitter.setStretchFactor(1, 2)  # Graphs

        main_layout.addWidget(main_splitter)  # Tasks below graphs
        
        # Add toolbar actions and Pomodoro timer
        tb = s.addToolBar('Actions')
        [tb.addAction(t,f) for t,f in [('Add Task',s.at),('Import Task',s.get_tasks),
                                        ('Status Report',s.sr)]]
        s.weekly_button = QPushButton("Weekly")
        s.monthly_button = QPushButton("Monthly")
        s.weekly_button.clicked.connect(lambda: update_plot("weekly"))
        s.monthly_button.clicked.connect(lambda: update_plot("monthly"))
        tb.addWidget(s.weekly_button)
        tb.addWidget(s.monthly_button)
        # Add Pomodoro timer to toolbar
        pomodoro_widget = QWidget()
        pomodoro_layout = QHBoxLayout(pomodoro_widget)
        s.timer_label = QLabel("25:00")
        s.timer_label.setStyleSheet("color: #00ffff; min-width: 60px;")
        start_button = QPushButton("?")
        start_button.setFixedSize(24, 24)
        start_button.clicked.connect(s.toggle_pomodoro)
        pomodoro_layout.addWidget(s.timer_label)
        pomodoro_layout.addWidget(start_button)
        pomodoro_layout.setContentsMargins(0, 0, 0, 0)
        tb.addWidget(pomodoro_widget)
        
        # Add completed tasks toggle
        s.toggle_completed = QAction('Show Completed Tasks', s)
        s.toggle_completed.setCheckable(True)
        s.toggle_completed.triggered.connect(s.toggle_completed_tasks)
        tb.addAction(s.toggle_completed)
        
                
        # Apply stylesheet
        s.setStyleSheet("""
            QMainWindow {
                background-color: #000810;
            }
            QToolBar {
                background-color: #001a1a;
                border: 1px solid #00ffff;
                border-radius: 5px;
                spacing: 3px;
            }
            QToolBar QToolButton {
                background-color: #002626;
                color: #00ffff;
                border: 1px solid #00ffff;
                border-radius: 3px;
                padding: 5px;
            }
            QDialog {
                background-color: #000810;
                color: #00ffff;
            }
            QDialog QWidget {
                background-color: #000810;
                color: #00ffff;
            }
            QToolBar QToolButton:hover {
                background-color: #004d4d;
            }
            QHeaderView::section {
                background-color: #001a1a;
                color: #00ffff;
                border: 1px solid #00ffff;
                padding: 4px;
            }
            QScrollBar {
                background-color: #001a1a;
                border: 1px solid #00ffff;
                border-radius: 4px;
            }
            QScrollBar::handle {
                background-color: #004d4d;
                border-radius: 3px;
            }
            QScrollBar::add-line, QScrollBar::sub-line {
                background: none;
            }
            QLabel {
                color: #00ffff;
            }
            QLineEdit, QSpinBox, QComboBox {
                background-color: #002626;
                color: #00ffff;
                border: 1px solid #00ffff;
                border-radius: 3px;
                padding: 3px;
            }
            QPushButton {
                background-color: #002626;
                color: #00ffff;
                border: 1px solid #00ffff;
                border-radius: 3px;
                padding: 5px 15px;
            }
            QPushButton:hover {
                background-color: #004d4d;
            }
            QTreeWidget {
                background-color: #000810;
                border: none;
            }
            QTreeWidget::item {
                background-color: #000810;
                color: #00ffff;
                border: none;
            }
            QProgressBar {
                background-color: #004d4d;
                height: 4px;
                border-radius: 2px;
                border: none;
            }
            QProgressBar::chunk {
                background-color: #00ffff;
                border-radius: 2px;
            }
        """)
        
        # Style all subplots
        for ax in s.a:
            ax.set_facecolor('black')
            if isinstance(ax, Axes3D):  # Check if it's a 3D subplot
                # Set pane colors to transparent
                ax.xaxis.pane.fill = False
                ax.yaxis.pane.fill = False
                ax.zaxis.pane.fill = False
                
                # Set pane edges to cyan
                ax.xaxis.pane.set_edgecolor('cyan')
                ax.yaxis.pane.set_edgecolor('cyan')
                ax.zaxis.pane.set_edgecolor('cyan')
                
                # Make panes transparent
                ax.xaxis.pane.set_alpha(0.0)
                ax.yaxis.pane.set_alpha(0.0)
                ax.zaxis.pane.set_alpha(0.0)
            ax.grid(False)
        
        # Set dark theme
        s.f.patch.set_facecolor('black')
        plt.style.use('dark_background')

        
        
        update_plot("weekly")
        s.c.draw()
        s.ul()
    def get_tasks(s):
        try:
            tasks = s.tasks_folder.Items
            print(f"Found {tasks.Count} tasks in Outlook")
            tasks.Sort("[DueDate]", True)
            task_list = [tasks[j] for j in range(min(50, tasks.Count))]
            
            d = QDialog(s)
            l = QVBoxLayout(d)
            lw = QListWidget()
            
            try:
                [lw.addItem(f"{t.Subject} - Due: {getattr(t,'DueDate','N/A')}") for t in task_list]
                if not lw.count():
                    QMessageBox.warning(s, "No Tasks", "No tasks found")
                    return
            except Exception as e:
                QMessageBox.warning(s, "Error", f"Task loading error: {str(e)}")
                return
        

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
                    print(f"Created task: {task.n} with status {task.st}")
                    s.ts+=[task];s.ul();s.uc()
                except Exception as e:
                    print(f"Detailed error: {str(e)}")
                    QMessageBox.warning(s,"Error",f"Task creation error: {str(e)}")
        except Exception as e:
            QMessageBox.warning(s, "Error", f"Outlook task retrieval error: {str(e)}")
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
        if i[4].text():  # If location is added
            s.locations_changed = True
        s.ul()
        s.uc()
        s.save_tasks()
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
            # Find and remove the task, including subtasks
            for task in s.ts:
                if task == t:  # Check if it's a top-level task
                    s.ts.remove(task)
                    break  # Stop searching once found
                else:
                    try:
                        task.s.remove(t)  # Try removing from subtasks
                        break  # Stop searching once found
                    except ValueError:
                        pass  # Continue if not found in subtasks

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
