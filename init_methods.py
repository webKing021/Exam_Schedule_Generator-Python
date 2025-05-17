def init_advanced_settings(self, frame):
    """Initialize the advanced settings section of the generator tab."""
    # Create a frame for settings with padding
    settings_frame = ttk.Frame(frame, padding=10)
    settings_frame.pack(fill=tk.BOTH, expand=True)
    
    # Configure grid
    settings_frame.columnconfigure(0, weight=0)
    settings_frame.columnconfigure(1, weight=1)
    
    # Gap between difficult exams
    ttk.Label(settings_frame, text="Exam Spacing Settings", font=("Helvetica", 12, "bold")).grid(
        row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
    
    ttk.Label(settings_frame, text="Gap between difficult exams (days):").grid(row=1, column=0, sticky=tk.W, pady=5)
    self.difficult_gap_var = tk.IntVar(value=2)
    ttk.Spinbox(settings_frame, from_=0, to=10, textvariable=self.difficult_gap_var, width=5).grid(
        row=1, column=1, sticky=tk.W, pady=5)
    
    ttk.Label(settings_frame, text="Gap between medium difficulty exams (days):").grid(row=2, column=0, sticky=tk.W, pady=5)
    self.medium_gap_var = tk.IntVar(value=1)
    ttk.Spinbox(settings_frame, from_=0, to=5, textvariable=self.medium_gap_var, width=5).grid(
        row=2, column=1, sticky=tk.W, pady=5)
    
    # Room allocation settings
    ttk.Separator(settings_frame, orient=tk.HORIZONTAL).grid(
        row=3, column=0, columnspan=2, sticky=tk.EW, pady=10)
    
    ttk.Label(settings_frame, text="Room Allocation Settings", font=("Helvetica", 12, "bold")).grid(
        row=4, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
    
    # Room selection
    ttk.Label(settings_frame, text="Select Rooms:").grid(row=5, column=0, sticky=tk.W, pady=5)
    
    # Create a frame for the room selection treeview
    room_frame = ttk.Frame(settings_frame)
    room_frame.grid(row=6, column=0, columnspan=2, sticky=tk.NSEW, pady=5)
    
    # Make the room frame expandable
    settings_frame.rowconfigure(6, weight=1)
    
    # Room selection treeview
    columns = ("id", "name", "type", "capacity", "selected")
    self.room_select_tree = ttk.Treeview(room_frame, columns=columns, show="headings", height=10)
    
    # Configure columns
    self.room_select_tree.heading("id", text="ID")
    self.room_select_tree.heading("name", text="Room Name")
    self.room_select_tree.heading("type", text="Type")
    self.room_select_tree.heading("capacity", text="Capacity")
    self.room_select_tree.heading("selected", text="Selected")
    
    # Set column widths
    self.room_select_tree.column("id", width=50)
    self.room_select_tree.column("name", width=150)
    self.room_select_tree.column("type", width=100)
    self.room_select_tree.column("capacity", width=80)
    self.room_select_tree.column("selected", width=80)
    
    # Add scrollbar
    room_scrollbar = ttk.Scrollbar(room_frame, orient=tk.VERTICAL, command=self.room_select_tree.yview)
    self.room_select_tree.configure(yscrollcommand=room_scrollbar.set)
    
    # Pack treeview and scrollbar
    self.room_select_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    room_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    # Bind click event to toggle selection
    self.room_select_tree.bind("<ButtonRelease-1>", self.toggle_room_selection)
    
    # Load rooms
    self.load_rooms_for_selection()

def init_preview_section(self, frame):
    """Initialize the preview section of the generator tab."""
    # Create a frame for the preview with padding
    preview_frame = ttk.Frame(frame, padding=10)
    preview_frame.pack(fill=tk.BOTH, expand=True)
    
    # Configure grid
    preview_frame.columnconfigure(0, weight=1)
    preview_frame.rowconfigure(1, weight=1)
    
    # Preview header
    ttk.Label(preview_frame, text="Schedule Preview", font=("Helvetica", 12, "bold")).grid(
        row=0, column=0, sticky=tk.W, pady=(0, 10))
    
    # Preview treeview
    preview_tree_frame = ttk.Frame(preview_frame)
    preview_tree_frame.grid(row=1, column=0, sticky=tk.NSEW)
    
    columns = ("date", "subject_code", "subject_name", "room", "time")
    self.preview_tree = ttk.Treeview(preview_tree_frame, columns=columns, show="headings", height=15)
    
    # Configure columns
    self.preview_tree.heading("date", text="Date")
    self.preview_tree.heading("subject_code", text="Subject Code")
    self.preview_tree.heading("subject_name", text="Subject Name")
    self.preview_tree.heading("room", text="Room")
    self.preview_tree.heading("time", text="Time")
    
    # Set column widths
    self.preview_tree.column("date", width=100)
    self.preview_tree.column("subject_code", width=100)
    self.preview_tree.column("subject_name", width=200)
    self.preview_tree.column("room", width=100)
    self.preview_tree.column("time", width=100)
    
    # Add scrollbars
    y_scrollbar = ttk.Scrollbar(preview_tree_frame, orient=tk.VERTICAL, command=self.preview_tree.yview)
    self.preview_tree.configure(yscrollcommand=y_scrollbar.set)
    
    x_scrollbar = ttk.Scrollbar(preview_tree_frame, orient=tk.HORIZONTAL, command=self.preview_tree.xview)
    self.preview_tree.configure(xscrollcommand=x_scrollbar.set)
    
    # Pack treeview and scrollbars
    self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
    
    # Preview controls
    controls_frame = ttk.Frame(preview_frame)
    controls_frame.grid(row=2, column=0, sticky=tk.EW, pady=10)
    
    # Generate preview button
    ttk.Button(controls_frame, text="Generate Preview", command=self.generate_preview).pack(side=tk.LEFT, padx=5)
    
    # Export buttons
    ttk.Button(controls_frame, text="Export to CSV", command=self.export_schedule_csv).pack(side=tk.LEFT, padx=5)
    ttk.Button(controls_frame, text="Export to PDF", command=self.export_schedule_pdf).pack(side=tk.LEFT, padx=5)

def load_rooms_for_selection(self):
    """Load rooms for selection in the advanced settings tab."""
    # Clear existing items
    for item in self.room_select_tree.get_children():
        self.room_select_tree.delete(item)
    
    try:
        # Ensure database connection is active
        if not self.ensure_connection():
            return
            
        # Fetch rooms from database - no parameters needed here
        try:
            print("Executing SELECT query for rooms in selection...")
            self.cursor.execute("SELECT id, name, type, capacity FROM rooms ORDER BY type, name")
            rooms = self.cursor.fetchall()
            if rooms is None:  # Handle None result
                rooms = []
            print(f"Found {len(rooms)} rooms for selection")
        except Error as e:
            print(f"Error fetching rooms for selection: {str(e)}")
            messagebox.showerror("Database Error", f"Failed to fetch rooms for selection: {str(e)}")
            return
        
        # Add rooms to treeview with selection status
        for room in rooms:
            try:
                self.room_select_tree.insert("", tk.END, values=(*room, "No"))
            except Exception as e:
                print(f"Error adding room to treeview: {str(e)}")
                continue
        
    except Error as e:
        # Log the error for debugging
        print(f"Error loading rooms for selection: {str(e)}")
        messagebox.showerror("Database Error", f"Failed to load rooms for selection: {str(e)}")
