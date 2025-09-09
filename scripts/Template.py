import clr
import System
from System.Runtime.InteropServices import Marshal
from AlibreScript.API import *
def printTraceBack():
    import traceback
    return
def show_error(msg, title='Error', include_trace=False):
    try:
        from System.Windows.Forms import MessageBox
        MessageBox.Show(str(msg), str(title), System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
    except:
        pass
    if include_trace:
        printTraceBack()
def show_info(msg, title='Info'):
    try:
        from System.Windows.Forms import MessageBox
        MessageBox.Show(str(msg), str(title), System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
    except:
        pass
def safe_try(fn):
    """Decorator-like wrapper for event handlers to avoid crashing the UI."""
    def _inner(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except Exception as ex:
            show_error('Unexpected error: %s' % ex, 'Unexpected Error', include_trace=True)
            return None
    return _inner
alibre = None
root = None
MyPart = None
try:
    alibre = Marshal.GetActiveObject("AlibreX.AutomationHook")
    root = alibre.Root
except Exception as ex:
    show_error('Could not connect to Alibre Automation. Details: %s' % ex, 'Alibre Connection Error', include_trace=True)
try:
    if root is not None:
        MyPart = Part(root.TopmostSession)
    else:
        MyPart = None
except Exception as ex:
    show_error('Could not get current Part session. Open a part and try again.\nDetails: %s' % ex, 'Part Error', include_trace=True)
    MyPart = None
try:
    clr.AddReference('System.Windows.Forms')
    clr.AddReference('System.Drawing')
except Exception as ex:
    show_error('Failed to load Windows Forms assemblies. Details: %s' % ex, 'Reference Load Error', include_trace=True)
from System.Windows.Forms import (
    ListBox, SelectionMode, Padding, AutoScaleMode,
    Timer, Button, ToolTip,
    DockStyle, Cursors, FlatStyle,
    Form, Label, CheckBox, NumericUpDown,
    MessageBox, Panel, TableLayoutPanel,
    FlowLayoutPanel, FlowDirection, RadioButton
)
from System.Drawing import Color, Size, SizeF, Font, FontStyle, SystemFonts
def scale_size(base_size, scale_factor=1.0):
    return int(base_size)
def get_professional_colors():
    return {
        'background': Color.FromArgb(250, 250, 250),
        'accent': Color.FromArgb(0, 122, 204),
        'accent_light': Color.FromArgb(230, 244, 255),
        'border': Color.FromArgb(204, 204, 204),
        'text': Color.FromArgb(64, 64, 64),
        'button_bg': Color.FromArgb(240, 240, 240)
    }
class SelectionListBox(ListBox):
    def __new__(cls):
        instance = ListBox.__new__(cls)
        try:
            instance.AutoScaleDimensions = SizeF(96, 96)
            instance.AutoScaleMode = AutoScaleMode.Dpi
            instance.IntegralHeight = 1
            instance.SelectionMode = SelectionMode.MultiExtended
            instance.BackColor = Color.White
            instance.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            instance.Font = Font(SystemFonts.DefaultFont.FontFamily, 9)
            import AlibreScript
            Root = AlibreScript.API.Global.Root
            instance.Root = Root
            instance.top_sess = instance.Root.TopmostSession
            instance.myTimer = Timer()
            instance.myTimer.Tick += instance.TimerEventProcessor
            instance.myTimer.Interval = 100
            instance.Enter += instance.onEnter_Selection
            instance.Leave += instance.onLeave_Selection
            instance.HandleDestroyed += instance.onHandleDestroyed
            instance.PreviousSelection = instance.Root.NewObjectCollector()
        except Exception as ex:
            show_error('Selection list init failed: %s' % ex, include_trace=True)
        return instance
    @safe_try
    def onEnter_Selection(self, sender, e):
        colors = get_professional_colors()
        sender.BackColor = colors['accent_light']
        sender.myTimer.Start()
    @safe_try
    def onLeave_Selection(self, sender, e):
        try:
            sender.myTimer.Stop()
        finally:
            sender.BackColor = Color.White
    @safe_try
    def onHandleDestroyed(self, sender, e):
        try:
            sender.myTimer.Stop()
        finally:
            pass
    @safe_try
    def TimerEventProcessor(self, sender, e):
        try:
            self.myTimer.Stop()
            if self.top_sess is None:
                return
            try:
                if self.PreviousSelection is None:
                    self.PreviousSelection = self.Root.NewObjectCollector()
            except:
                return
            NewSelections = getattr(self.top_sess, 'SelectedObjects', None)
            if NewSelections is None:
                return
            try:
                count = int(NewSelections.Count)
            except:
                return
            for a in range(0, count):
                item = NewSelections.Item(a)
                tgt = getattr(item, 'Target', None)
                tname = ''
                try:
                    tname = str(tgt.GetType().Name)
                except:
                    try:
                        tname = str(tgt.Type)
                    except:
                        tname = ''
                if tgt is not None and 'SKETCH' in tname.upper():
                    try:
                        sketch_name = str(tgt.Name)
                    except:
                        sketch_name = str(getattr(item, 'DisplayName', 'Sketch'))
                    if self.Items.Count == 0 or sketch_name != self.Items[0]:
                        self.Items.Clear()
                        try:
                            self.PreviousSelection.Clear()
                        except:
                            try:
                                self.PreviousSelection = self.Root.NewObjectCollector()
                            except:
                                pass
                        self.Items.Add(sketch_name)
                        try:
                            self.PreviousSelection.Add(item)
                        except:
                            pass
                    break
        finally:
            try:
                self.myTimer.Start()
            except:
                pass
def create_professional_button(text, is_primary=False):
    colors = get_professional_colors()
    btn = Button()
    btn.Text = text
    btn.FlatStyle = FlatStyle.Flat
    btn.Font = Font(SystemFonts.DefaultFont.FontFamily, 9, FontStyle.Regular)
    btn.UseVisualStyleBackColor = False
    if is_primary:
        btn.BackColor = colors['accent']
        btn.ForeColor = Color.White
        btn.FlatAppearance.BorderColor = colors['accent']
    else:
        btn.BackColor = colors['button_bg']
        btn.ForeColor = colors['text']
        btn.FlatAppearance.BorderColor = colors['border']
    btn.FlatAppearance.BorderSize = 1
    btn.Cursor = Cursors.Hand
    return btn
def create_professional_label(text, is_header=False):
    colors = get_professional_colors()
    lbl = Label()
    lbl.Text = text
    lbl.ForeColor = colors['text']
    lbl.AutoSize = True
    lbl.TextAlign = System.Drawing.ContentAlignment.TopLeft
    lbl.Font = Font(SystemFonts.DefaultFont.FontFamily, 10 if is_header else 9,
                    FontStyle.Bold if is_header else FontStyle.Regular)
    return lbl
def create_professional_checkbox(text):
    colors = get_professional_colors()
    chk = CheckBox()
    chk.Text = text
    chk.ForeColor = colors['text']
    chk.Font = Font(SystemFonts.DefaultFont.FontFamily, 9)
    chk.UseVisualStyleBackColor = True
    chk.AutoSize = True
    chk.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    return chk
def create_professional_numericupdown():
    colors = get_professional_colors()
    num = NumericUpDown()
    num.BackColor = Color.White
    num.ForeColor = colors['text']
    num.Font = Font(SystemFonts.DefaultFont.FontFamily, scale_size(9))
    num.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    num.Margin = Padding(2)
    return num
def create_professional_radiobutton(text):
    colors = get_professional_colors()
    rb = RadioButton()
    rb.Text = text
    rb.AutoSize = True
    rb.Font = Font(SystemFonts.DefaultFont.FontFamily, 9)
    rb.ForeColor = colors['text']
    rb.Margin = Padding(0, 0, 16, 0)
    return rb
def show_form():
    if MyPart is None:
        show_error('No active Part session was found. Open a part and run the script again.', 'No Part Session')
        return None
    colors = get_professional_colors()
    form = Form()
    form.Text = 'Advanced Sweep Tool'
    form.AutoSize = False
    form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    form.MaximizeBox = False
    form.MinimizeBox = False
    form.ShowInTaskbar = False
    form.ShowIcon = False
    form.TopMost = True
    form.BackColor = colors['background']
    form.Font = Font(SystemFonts.DefaultFont.FontFamily, 9)
    form.Padding = Padding(12)
    main_panel = Panel()
    main_panel.Dock = DockStyle.Fill
    main_panel.BackColor = colors['background']
    main_panel.Padding = Padding(8)
    main_panel.AutoSize = True
    main_panel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
    form.Controls.Add(main_panel)
    table = TableLayoutPanel()
    table.ColumnCount = 1
    table.RowCount = 0
    table.Dock = DockStyle.Fill
    table.AutoSize = True
    table.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
    table.Padding = Padding(0)
    table.Margin = Padding(0)
    table.ColumnStyles.Add(System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0))
    main_panel.Controls.Add(table)
    control_spacing = 8
    section_spacing = 16
    def add_control_row(ctrl, extra_margin_bottom=None, fixed_height=None):
        table.RowCount += 1
        if fixed_height is not None:
            table.RowStyles.Add(System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, fixed_height))
            ctrl.Height = fixed_height
        else:
            table.RowStyles.Add(System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        mb = control_spacing if extra_margin_bottom is None else extra_margin_bottom
        ctrl.Margin = Padding(0, 0, 0, mb)
        ctrl.Dock = DockStyle.Fill
        table.Controls.Add(ctrl, 0, table.RowCount - 1)
    add_control_row(create_professional_label("Path Sketch Selection", True), extra_margin_bottom=6)
    add_control_row(create_professional_label("Select a 2D or 3D sketch to use as the sweep path:"))
    sel_sketch = SelectionListBox()
    sel_sketch.IntegralHeight = False
    sel_sketch.Height = 90
    add_control_row(sel_sketch, fixed_height=90)
    chk_3d = create_professional_checkbox("Use 3D Path")
    add_control_row(chk_3d, extra_margin_bottom=section_spacing)
    add_control_row(create_professional_label("Units", True), extra_margin_bottom=6)
    add_control_row(create_professional_label("Select measurement units:"))
    units_flow = FlowLayoutPanel()
    units_flow.FlowDirection = FlowDirection.LeftToRight
    units_flow.WrapContents = True
    units_flow.AutoSize = True
    rb_mm = create_professional_radiobutton("Millimeters")
    rb_in = create_professional_radiobutton("Inches")
    rb_cm = create_professional_radiobutton("Centimeters")
    rb_mm.Checked = True
    units_flow.Controls.Add(rb_mm)
    units_flow.Controls.Add(rb_in)
    units_flow.Controls.Add(rb_cm)
    add_control_row(units_flow, extra_margin_bottom=section_spacing)
    add_control_row(create_professional_label("Profile Configuration", True), extra_margin_bottom=6)
    add_control_row(create_professional_label("Profile Type:"))
    prof_flow = FlowLayoutPanel()
    prof_flow.FlowDirection = FlowDirection.LeftToRight
    prof_flow.WrapContents = True
    prof_flow.AutoSize = True
    rb_circle = create_professional_radiobutton("Circle")
    rb_rect = create_professional_radiobutton("Rectangle")
    rb_circle.Checked = True
    prof_flow.Controls.Add(rb_circle)
    prof_flow.Controls.Add(rb_rect)
    add_control_row(prof_flow)
    lbl_size = create_professional_label("Profile Size:")
    add_control_row(lbl_size)
    num_size = create_professional_numericupdown()
    num_size.Value = 50
    num_size.DecimalPlaces = 2
    num_size.Maximum = 10000
    num_size.Minimum = 0.01
    num_size.Height = 28
    add_control_row(num_size, extra_margin_bottom=section_spacing)
    add_control_row(create_professional_label("Hollow Options", True), extra_margin_bottom=6)
    chk_hollow = create_professional_checkbox("Create hollow profile")
    add_control_row(chk_hollow)
    lbl_thick = create_professional_label("Wall Thickness:")
    add_control_row(lbl_thick)
    num_thick = create_professional_numericupdown()
    num_thick.Value = 2
    num_thick.DecimalPlaces = 2
    num_thick.Maximum = 1000
    num_thick.Minimum = 0.0
    num_thick.Height = 28
    num_thick.Enabled = False
    add_control_row(num_thick, extra_margin_bottom=section_spacing)
    def chk_changed(sender, e):
        num_thick.Enabled = chk_hollow.Checked
        lbl_thick.ForeColor = colors['text'] if chk_hollow.Checked else colors['border']
    chk_hollow.CheckedChanged += safe_try(chk_changed)
    def current_unit_text():
        if rb_mm.Checked:
            return "Millimeters"
        if rb_in.Checked:
            return "Inches"
        if rb_cm.Checked:
            return "Centimeters"
        return "Millimeters"
    def update_unit_labels(sender, e):
        try:
            unit_map = {"Millimeters": "mm", "Inches": "in", "Centimeters": "cm"}
            unit_text = current_unit_text()
            unit_abbr = unit_map.get(unit_text, "units")
            lbl_size.Text = "Profile Size (%s):" % unit_abbr
            lbl_thick.Text = "Wall Thickness (%s):" % unit_abbr
        except Exception as ex:
            print('Unit label update failed: %s' % ex)
    rb_mm.CheckedChanged += safe_try(update_unit_labels)
    rb_in.CheckedChanged += safe_try(update_unit_labels)
    rb_cm.CheckedChanged += safe_try(update_unit_labels)
    update_unit_labels(None, None)
    chk_stay_open = create_professional_checkbox("Stay open after creating")
    add_control_row(chk_stay_open, extra_margin_bottom=8)
    def figures_count(figs):
        try:
            return int(figs.Count)
        except:
            try:
                return len(figs)
            except:
                c = 0
                try:
                    for _ in figs:
                        c += 1
                except:
                    pass
                return c
    def figure_at(figs, idx):
        try:
            return figs.Item(idx)
        except:
            try:
                return figs[idx]
            except:
                i = 0
                try:
                    for f in figs:
                        if i == idx:
                            return f
                        i += 1
                except:
                    return None
        return None
    def first_curve_figure(figs):
        n = figures_count(figs)
        for i in range(0, n):
            f = figure_at(figs, i)
            if f is None:
                continue
            try:
                if hasattr(f, 'IsReference') and f.IsReference:
                    continue
                t = f.GetType().Name.upper()
                if 'POINT' in t:
                    continue
                if hasattr(f, 'GetPointAt') or (hasattr(f, 'StartPoint') and hasattr(f, 'EndPoint')):
                    return f
            except:
                continue
        return None
    def clean_name(nm):
        if nm is None:
            return None
        s = str(nm)
        return s.split(':')[-1].strip() if ':' in s else s
    def close_form_safely():
        try:
            if hasattr(sel_sketch, 'myTimer') and sel_sketch.myTimer is not None:
                sel_sketch.myTimer.Stop()
                sel_sketch.myTimer.Dispose()
        except:
            pass
        try:
            form.Close()
            form.Dispose()
        except:
            pass
    def reset_tool_state():
        try:
            rb_mm.Checked = True
            rb_in.Checked = False
            rb_cm.Checked = False
            update_unit_labels(None, None)
            rb_circle.Checked = True
            rb_rect.Checked = False
            try:
                num_size.Value = 50
            except:
                pass
            chk_hollow.Checked = False
            try:
                num_thick.Value = 2
            except:
                pass
            num_thick.Enabled = False
            lbl_thick.ForeColor = colors['border']
            chk_3d.Checked = False
            try:
                sel_sketch.Items.Clear()
            except:
                pass
            try:
                if getattr(sel_sketch, 'Root', None) is not None:
                    sel_sketch.PreviousSelection = sel_sketch.Root.NewObjectCollector()
                else:
                    sel_sketch.PreviousSelection = None
            except:
                sel_sketch.PreviousSelection = None
            try:
                sel_sketch.myTimer.Stop()
                sel_sketch.myTimer.Start()
            except:
                pass
            try:
                sel_sketch.Focus()
            except:
                pass
        except:
            pass
    @safe_try
    def ok_click(sender, e):
        try:
            if rb_mm.Checked:
                Units.Current = UnitTypes.Millimeters
            elif rb_in.Checked:
                Units.Current = UnitTypes.Inches
            elif rb_cm.Checked:
                Units.Current = UnitTypes.Centimeters
            else:
                show_error('Please select a valid unit.', 'Invalid Units')
                return
        except Exception as ex:
            show_error('Failed to set units: %s' % ex, include_trace=True)
            return
        Use3DPath = chk_3d.Checked
        if rb_circle.Checked:
            ProfileTypeIndex = 0
        elif rb_rect.Checked:
            ProfileTypeIndex = 1
        else:
            ProfileTypeIndex = -1
        try:
            ProfileSize = float(num_size.Value)
            if ProfileSize <= 0:
                show_error('Profile size must be > 0.', 'Invalid Size')
                return
        except Exception as ex:
            show_error('Invalid profile size: %s' % ex)
            return
        try:
            IsHollow = chk_hollow.Checked
            Thickness = float(num_thick.Value)
            if IsHollow:
                if Thickness <= 0:
                    show_error('Wall thickness must be > 0.', 'Invalid Thickness')
                    return
                if ProfileTypeIndex in (0, 1) and Thickness >= (ProfileSize / 2.0):
                    show_error('Wall thickness must be < half the profile size.', 'Invalid Thickness')
                    return
        except Exception as ex:
            show_error('Invalid thickness value: %s' % ex)
            return
        if sel_sketch.Items.Count == 0 or sel_sketch.PreviousSelection is None:
            show_info("Please select a path sketch", "Input Required")
            return
        try:
            selected_item = sel_sketch.PreviousSelection.Item(0)
        except Exception as ex:
            show_error('Could not read selected sketch: %s' % ex, include_trace=True)
            return
        target = getattr(selected_item, 'Target', None)
        try:
            sketch_name = clean_name(getattr(target, 'Name', None))
        except:
            sketch_name = None
        if sketch_name is None:
            try:
                sketch_name = clean_name(sel_sketch.Items[0])
            except:
                sketch_name = None
        if sketch_name is None:
            show_error('Could not resolve the selected sketch name.', 'Sketch Error')
            return
        PathSketch = None
        try:
            if Use3DPath:
                try:
                    PathSketch = MyPart.Get3DSketch(sketch_name)
                except Exception:
                    PathSketch = MyPart.Get3DSketch(str(sel_sketch.Items[0]))
            else:
                try:
                    PathSketch = MyPart.GetSketch(sketch_name)
                except Exception:
                    PathSketch = MyPart.GetSketch(str(sel_sketch.Items[0]))
        except Exception as ex:
            show_error('Use 3D Path must match the selected sketch type. Failed getting "%s". Details: %s' %
                       (sketch_name, ex), '2D vs 3D Sketch Error', include_trace=True)
            return
        if PathSketch is None:
            show_error("A Path Sketch must be selected.", "Missing Path Sketch")
            return
        try:
            is3d = '3D' in PathSketch.GetType().Name
            if is3d and not Use3DPath:
                show_error('Selected sketch is 3D. Check "Use 3D Path".', '3D Path Required')
                return
            if (not is3d) and Use3DPath:
                show_error('Selected sketch is 2D. Uncheck "Use 3D Path" or pick a 3D sketch.', '2D Path Selected')
                return
        except:
            pass
        FirstFigure = first_curve_figure(PathSketch.Figures)
        if FirstFigure is None:
            show_error("The path sketch must contain a valid non-reference curve.", "Invalid Path Geometry")
            return
        try:
            def as_xyz(p):
                try:
                    return [float(p.X), float(p.Y), float(p.Z)]
                except:
                    pass
                try:
                    n = len(p)
                    if n == 3:
                        return [float(p[0]), float(p[1]), float(p[2])]
                    if n == 2:
                        return [float(p[0]), float(p[1]), 0.0]
                except:
                    pass
                raise Exception("Unsupported point type for xyz conversion: %r" % (type(p),))
            def list_features(part):
                feats = []
                try:
                    col = part.Features
                    try:
                        count = int(col.Count)
                        for i in range(count):
                            try:
                                feats.append(col.Item(i))
                            except:
                                pass
                    except:
                        for f in col:
                            feats.append(f)
                except:
                    pass
                return feats
            def feat_key(f):
                try:
                    h = int(f.GetHashCode())
                except:
                    h = None
                try:
                    n = str(f.Name)
                except:
                    n = ""
                try:
                    t = str(f.GetType().Name)
                except:
                    t = ""
                return (h, n, t)
            if hasattr(FirstFigure, 'GetPointAt'):
                if is3d:
                    sp = as_xyz(FirstFigure.GetPointAt(0.0))
                    np = as_xyz(FirstFigure.GetPointAt(0.01))
                else:
                    sp2 = FirstFigure.GetPointAt(0.0)
                    np2 = FirstFigure.GetPointAt(0.01)
                    sp = as_xyz(PathSketch.PointtoGlobal(sp2[0], sp2[1]))
                    np = as_xyz(PathSketch.PointtoGlobal(np2[0], np2[1]))
            else:
                if is3d:
                    sp = as_xyz(FirstFigure.StartPoint)
                    np = as_xyz(FirstFigure.EndPoint)
                else:
                    sp2 = FirstFigure.StartPoint
                    np2 = FirstFigure.EndPoint
                    sp = as_xyz(PathSketch.PointtoGlobal(sp2[0], sp2[1]))
                    np = as_xyz(PathSketch.PointtoGlobal(np2[0], np2[1]))
            dir_vec = [np[i] - sp[i] for i in range(3)]
            created_plane = None
            created_sketch = None
            created_sweep = None
            pre_feats = list_features(MyPart)
            pre_keys = set(feat_key(f) for f in pre_feats)
            try:
                MyPart.PauseUpdating()
            except:
                pass
            try:
                created_plane = MyPart.AddPlane("SweepProfilePlane", dir_vec, sp)
                created_sketch = MyPart.AddSketch("SweepProfile", created_plane)
                pt2d = created_sketch.GlobaltoPoint(sp[0], sp[1], sp[2])
                try:
                    cx, cy = float(pt2d.X), float(pt2d.Y)
                except:
                    cx, cy = float(pt2d[0]), float(pt2d[1])
                if rb_circle.Checked:
                    created_sketch.AddCircle(cx, cy, ProfileSize / 2.0, False)
                elif rb_rect.Checked:
                    half = ProfileSize / 2.0
                    created_sketch.AddRectangle(cx - half, cy - half, cx + half, cy + half, False)
                else:
                    raise Exception("Unsupported profile type.")
                if IsHollow:
                    if rb_circle.Checked:
                        inner_r = (ProfileSize / 2.0) - Thickness
                        if inner_r <= 0:
                            raise Exception("Wall thickness too large for the chosen size.")
                        created_sketch.AddCircle(cx, cy, inner_r, False)
                    elif rb_rect.Checked:
                        inner_h = (ProfileSize / 2.0) - Thickness
                        if inner_h <= 0:
                            raise Exception("Wall thickness too large for the chosen size.")
                        created_sketch.AddRectangle(cx - inner_h, cy - inner_h, cx + inner_h, cy + inner_h, False)
                created_sweep = MyPart.AddSweepBoss(
                    "AdvancedSweep",
                    created_sketch,
                    PathSketch,
                    False,
                    MyPart.EndCondition.EntirePath,
                    None, 0, 0, False
                )
            except Exception as build_ex:
                try:
                    post_feats = list_features(MyPart)
                    for f in post_feats:
                        if feat_key(f) not in pre_keys:
                            try:
                                MyPart.RemoveFeature(f)
                            except:
                                pass
                except:
                    pass
                try:
                    if created_sketch is not None:
                        MyPart.RemoveSketch(created_sketch)
                except:
                    pass
                try:
                    if created_plane is not None:
                        MyPart.RemovePlane(created_plane)
                except:
                    pass
                try:
                    if created_sweep is not None:
                        MyPart.RemoveFeature(created_sweep)
                except:
                    pass
                raise build_ex
            finally:
                try:
                    MyPart.ResumeUpdating()
                except:
                    pass
            if created_sweep is None:
                try:
                    MyPart.PauseUpdating()
                except:
                    pass
                try:
                    post_feats = list_features(MyPart)
                    for f in post_feats:
                        if feat_key(f) not in pre_keys:
                            try:
                                MyPart.RemoveFeature(f)
                            except:
                                pass
                except:
                    pass
                try:
                    if created_sketch is not None:
                        MyPart.RemoveSketch(created_sketch)
                except:
                    pass
                try:
                    if created_plane is not None:
                        MyPart.RemovePlane(created_plane)
                except:
                    pass
                finally:
                    try:
                        MyPart.ResumeUpdating()
                    except:
                        pass
                raise Exception("Sweep failed to create for an unknown reason.")
            if not chk_stay_open.Checked:
                close_form_safely()
            else:
                reset_tool_state()
                try:
                    form.Activate()
                except:
                    pass
            return
        except Exception as ex:
            show_error("Failed to create sweep: %s" % ex, "Create Sweep Error", include_trace=True)
            return
    def cancel_click(sender, e):
        close_form_safely()
    btn_ok = create_professional_button("Advanced Sweep Tool", True)
    btn_ok.Dock = DockStyle.Fill
    add_control_row(btn_ok, extra_margin_bottom=8, fixed_height=60)
    btn_cancel = create_professional_button("Close", False)
    btn_cancel.Dock = DockStyle.Fill
    add_control_row(btn_cancel, extra_margin_bottom=0, fixed_height=50)
    btn_ok.Click += ok_click
    btn_cancel.Click += safe_try(cancel_click)
    tooltip = ToolTip()
    tooltip.SetToolTip(sel_sketch, "Click here, then select a sketch in the Alibre workspace")
    tooltip.SetToolTip(chk_3d, "Check this if you selected a 3D sketch")
    tooltip.SetToolTip(rb_circle, "Use a circular cross-section")
    tooltip.SetToolTip(rb_rect, "Use a rectangular cross-section")
    tooltip.SetToolTip(num_size, "Set the outer dimension of the profile")
    tooltip.SetToolTip(chk_hollow, "Create a hollow tube instead of solid")
    tooltip.SetToolTip(num_thick, "Wall thickness for hollow profiles")
    tooltip.SetToolTip(rb_mm, "Use millimeters")
    tooltip.SetToolTip(rb_in, "Use inches")
    tooltip.SetToolTip(rb_cm, "Use centimeters")
    tooltip.SetToolTip(chk_stay_open, "Leave this window open after creating the sweep")
    table.PerformLayout()
    main_panel.PerformLayout()
    try:
        pref = table.PreferredSize
        total_hpad = form.Padding.Left + form.Padding.Right + main_panel.Padding.Left + main_panel.Padding.Right
        total_vpad = form.Padding.Top + form.Padding.Bottom + main_panel.Padding.Top + main_panel.Padding.Bottom
        min_width = 460
        width = max(min_width, pref.Width + total_hpad)
        height = pref.Height + total_vpad
        height = min(height, 1000)
        form.ClientSize = Size(int(width), int(height))
        form.MinimumSize = Size(min_width, 240)
    except Exception as ex:
        print('Form sizing warning: %s' % ex)
    try:
        form.Show()
    except Exception as ex:
        show_error('Failed to display form: %s' % ex, include_trace=True)
        return None
    return form
template_form = None
try:
    template_form = show_form()
except Exception as ex:
    show_error('Fatal error while creating the UI: %s' % ex, include_trace=True)
