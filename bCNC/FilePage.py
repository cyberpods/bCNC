# $Id$
#
# Author: vvlachoudis@gmail.com
# Date: 18-Jun-2015

import os
import sys
from tkinter import (
    YES,
    W,
    E,
    EW,
    NSEW,
    BOTH,
    LEFT,
    TOP,
    RIGHT,
    BooleanVar,
    Checkbutton,
    Label,
    Menu,
    Toplevel,
    Listbox,
    Frame,
    Scrollbar,
    Button,
    messagebox,
    StringVar,
    Entry,
)
import CNCRibbon
import Ribbon
import tkExtra
import Utils

from Helpers import N_

__author__ = "Vasilis Vlachoudis"
__email__ = "vvlachoudis@gmail.com"

try:
    from serial.tools.list_ports import comports
except Exception:
    print("Using fallback Utils.comports()!")
    from Utils import comports

BAUDS = [2400, 4800, 9600, 19200, 38400, 57600, 115200, 230400]


# =============================================================================
# Recent Menu button
# =============================================================================


class _RecentMenuButton(Ribbon.MenuButton):
    # ----------------------------------------------------------------------
    def createMenu(self):
        menu = Menu(self, tearoff=0, activebackground=Ribbon._ACTIVE_COLOR)
        for i in range(Utils._maxRecent):
            filename = Utils.getRecent(i)
            if filename is None:
                break
            path = os.path.dirname(filename)
            fn = os.path.basename(filename)
            menu.add_command(
                label="%d %s" % (i + 1, fn),
                compound=LEFT,
                image=Utils.icons["new"],
                accelerator=path,  # Show as accelerator in order to be aligned
                command=lambda s=self, i=i: s.event_generate(
                    "<<Recent%d>>" % (i)),
            )
        if i == 0:  # no entry
            self.event_generate("<<Open>>")
            return None
        return menu


# =============================================================================
# File Group
# =============================================================================
class FileGroup(CNCRibbon.ButtonGroup):
    def __init__(self, master, app):
        CNCRibbon.ButtonGroup.__init__(self, master, N_("File"), app)
        self.grid3rows()

        # ---
        col, row = 0, 0
        b = Ribbon.LabelButton(
            self.frame,
            self,
            "<<New>>",
            image=Utils.icons["new32"],
            text=_("New"),
            compound=TOP,
            background=Ribbon._BACKGROUND,
        )
        b.grid(row=row, column=col, rowspan=3, padx=0, pady=0, sticky=NSEW)
        tkExtra.Balloon.set(b, _("New gcode/dxf file"))
        self.addWidget(b)

        # ---
        col, row = 1, 0
        b = Ribbon.LabelButton(
            self.frame,
            self,
            "<<Open>>",
            image=Utils.icons["open32"],
            background=Ribbon._BACKGROUND,
        )
        b.grid(row=row, column=col, rowspan=2, padx=0, pady=0, sticky=NSEW)
        tkExtra.Balloon.set(b, _("Open existing gcode/dxf file [Ctrl-O]"))
        self.addWidget(b)

        col, row = 1, 2
        b = _RecentMenuButton(
            self.frame,
            None,
            text=_("Open"),
            image=Utils.icons["triangle_down"],
            compound=RIGHT,
            background=Ribbon._BACKGROUND,
        )
        b.grid(row=row, column=col, padx=0, pady=0, sticky=NSEW)
        tkExtra.Balloon.set(b, _("Open recent file"))
        self.addWidget(b)

        # --- Favorites Button ---
        col, row = 4, 0
        # Create FavoritesButton that directly calls openFavoriteMenu when clicked
        b = Ribbon.LabelButton(
            self.frame,
            image=Utils.icons["open32"],
            text=_("Favorites"),
            compound=TOP,
            command=self.openFavoriteMenu,
            background=Ribbon._BACKGROUND,
        )
        b.grid(row=row, column=col, rowspan=3, padx=0, pady=0, sticky=NSEW)
        tkExtra.Balloon.set(b, _("Access and manage favorite files"))
        self.addWidget(b)

        # ---
        col, row = 2, 0
        b = Ribbon.LabelButton(
            self.frame,
            self,
            "<<Import>>",
            image=Utils.icons["import32"],
            text=_("Import"),
            compound=TOP,
            background=Ribbon._BACKGROUND,
        )
        b.grid(row=row, column=col, rowspan=3, padx=0, pady=0, sticky=NSEW)
        tkExtra.Balloon.set(b, _("Import gcode/dxf file"))
        self.addWidget(b)

        # ---
        col, row = 3, 0
        b = Ribbon.LabelButton(
            self.frame,
            self,
            "<<Save>>",
            image=Utils.icons["save32"],
            command=app.save,
            background=Ribbon._BACKGROUND,
        )
        b.grid(row=row, column=col, rowspan=2, padx=0, pady=0, sticky=NSEW)
        tkExtra.Balloon.set(b, _("Save gcode/dxf file [Ctrl-S]"))
        self.addWidget(b)

        col, row = 3, 2
        b = Ribbon.LabelButton(
            self.frame,
            self,
            "<<SaveAs>>",
            text=_("Save"),
            image=Utils.icons["triangle_down"],
            compound=RIGHT,
            background=Ribbon._BACKGROUND,
        )
        b.grid(row=row, column=col, padx=0, pady=0, sticky=NSEW)
        tkExtra.Balloon.set(b, _("Save gcode/dxf AS"))
        self.addWidget(b)

    # ----------------------------------------------------------------------
    def openFavoriteMenu(self):
        """Show a menu with favorites"""
        # Load favorites
        favorites = []
        for i in range(20):  # Maximum of 20 favorites
            favorite = Utils.getStr("Favorites", f"favorite.{i}", "")
            if favorite:
                favorites.append(favorite)
            else:
                break

        # Create popup menu
        menu = Menu(self.master, tearoff=0)

        # Add management options
        menu.add_command(
            label=_("Add Current File to Favorites"),
            command=self.addFavorite
        )

        menu.add_command(
            label=_("Manage Favorites..."),
            command=self.manageFavorites
        )

        if favorites:
            menu.add_separator()

            # Add favorites
            for filename in favorites:
                fn = os.path.basename(filename)
                menu.add_command(
                    label=fn,
                    command=lambda f=filename: self.openFavorite(f)
                )

        # Display menu at mouse position
        menu.tk_popup(self.winfo_pointerx(), self.winfo_pointery())

    # ----------------------------------------------------------------------
    def addFavorite(self):
        """Add current file to favorites"""
        filename = self.app.gcode.filename
        if not filename:
            messagebox.showinfo(
                _("No File Open"),
                _("Please open a file before adding it to favorites.")
            )
            return

        # Load existing favorites
        favorites = []
        for i in range(20):  # Maximum of 20 favorites
            favorite = Utils.getStr("Favorites", f"favorite.{i}", "")
            if favorite:
                favorites.append(favorite)
            else:
                break

        if filename in favorites:
            messagebox.showinfo(
                _("Already in Favorites"),
                _("'{}' is already in your favorites.").format(os.path.basename(filename))
            )
            return

        favorites.append(filename)

        # Save favorites
        if "Favorites" not in Utils.config.sections():
            Utils.config.add_section("Favorites")
        else:
            # Clear existing favorites but keep the section
            for option in Utils.config.options("Favorites"):
                Utils.config.remove_option("Favorites", option)

        for i, favorite in enumerate(favorites):
            Utils.setStr("Favorites", f"favorite.{i}", favorite)

        messagebox.showinfo(
            _("Added to Favorites"),
            _("'{}' has been added to your favorites.").format(os.path.basename(filename))
        )

    # ----------------------------------------------------------------------
    def openFavorite(self, filename):
        """Open a favorite file"""
        if os.path.exists(filename):
            self.app.load(filename)
        else:
            messagebox.showerror(
                _("Error"),
                _("File not found: {}").format(filename)
            )
            # Ask if the user wants to remove from favorites
            if messagebox.askyesno(
                    _("Remove Favorite"),
                    _("File no longer exists. Remove from favorites?")
            ):
                # Load existing favorites
                favorites = []
                for i in range(20):
                    favorite = Utils.getStr("Favorites", f"favorite.{i}", "")
                    if favorite:
                        favorites.append(favorite)
                    else:
                        break

                if filename in favorites:
                    favorites.remove(filename)

                    # Save favorites
                    if "Favorites" not in Utils.config.sections():
                        Utils.config.add_section("Favorites")
                    else:
                        # Clear existing favorites but keep the section
                        for option in Utils.config.options("Favorites"):
                            Utils.config.remove_option("Favorites", option)

                    for i, favorite in enumerate(favorites):
                        Utils.setStr("Favorites", f"favorite.{i}", favorite)

    # ----------------------------------------------------------------------
    def manageFavorites(self):
        """Open a dialog to manage favorites"""
        # Load favorites
        favorites = []
        for i in range(20):
            favorite = Utils.getStr("Favorites", f"favorite.{i}", "")
            if favorite:
                favorites.append(favorite)
            else:
                break

        # Create dialog
        dialog = Toplevel(self.master)
        dialog.title(_("Manage Favorites"))
        dialog.transient(self.master)
        dialog.resizable(width=True, height=True)
        dialog.minsize(400, 300)
        dialog.grab_set()

        # List frame
        frame = Frame(dialog)
        frame.pack(fill=BOTH, expand=YES, padx=5, pady=5)

        scrollbar = Scrollbar(frame)
        scrollbar.pack(side=RIGHT, fill="y")

        listbox = Listbox(frame, yscrollcommand=scrollbar.set)
        listbox.pack(side=LEFT, fill=BOTH, expand=YES)

        scrollbar.config(command=listbox.yview)

        # Populate listbox
        for favorite in favorites:
            listbox.insert("end", favorite)

        # Button frame
        button_frame = Frame(dialog)
        button_frame.pack(fill="x", padx=5, pady=5)

        def remove_favorite():
            selection = listbox.curselection()
            if not selection:
                return

            index = selection[0]
            path = listbox.get(index)

            if messagebox.askyesno(
                    _("Confirm Removal"),
                    _("Remove '{}' from favorites?").format(os.path.basename(path))
            ):
                favorites.pop(index)
                listbox.delete(index)

                # Save favorites
                if "Favorites" not in Utils.config.sections():
                    Utils.config.add_section("Favorites")
                else:
                    # Clear existing favorites but keep the section
                    for option in Utils.config.options("Favorites"):
                        Utils.config.remove_option("Favorites", option)

                for i, favorite in enumerate(favorites):
                    Utils.setStr("Favorites", f"favorite.{i}", favorite)

        def move_up():
            selection = listbox.curselection()
            if not selection or selection[0] == 0:
                return

            index = selection[0]
            item = favorites[index]
            favorites.pop(index)
            favorites.insert(index - 1, item)

            listbox.delete(0, "end")
            for favorite in favorites:
                listbox.insert("end", favorite)

            listbox.selection_set(index - 1)

            # Save favorites
            if "Favorites" not in Utils.config.sections():
                Utils.config.add_section("Favorites")
            else:
                # Clear existing favorites but keep the section
                for option in Utils.config.options("Favorites"):
                    Utils.config.remove_option("Favorites", option)

            for i, favorite in enumerate(favorites):
                Utils.setStr("Favorites", f"favorite.{i}", favorite)

        def move_down():
            selection = listbox.curselection()
            if not selection or selection[0] == len(favorites) - 1:
                return

            index = selection[0]
            item = favorites[index]
            favorites.pop(index)
            favorites.insert(index + 1, item)

            listbox.delete(0, "end")
            for favorite in favorites:
                listbox.insert("end", favorite)

            listbox.selection_set(index + 1)

            # Save favorites
            if "Favorites" not in Utils.config.sections():
                Utils.config.add_section("Favorites")
            else:
                # Clear existing favorites but keep the section
                for option in Utils.config.options("Favorites"):
                    Utils.config.remove_option("Favorites", option)

            for i, favorite in enumerate(favorites):
                Utils.setStr("Favorites", f"favorite.{i}", favorite)

        # Add favorite manually
        def add_favorite():
            filename = Utils.getOpenFilename(self.master)
            if filename:
                if filename in favorites:
                    messagebox.showinfo(
                        _("Already in Favorites"),
                        _("'{}' is already in your favorites.").format(os.path.basename(filename))
                    )
                    return

                favorites.append(filename)
                listbox.insert("end", filename)

                # Save favorites
                if "Favorites" not in Utils.config.sections():
                    Utils.config.add_section("Favorites")
                else:
                    # Clear existing favorites but keep the section
                    for option in Utils.config.options("Favorites"):
                        Utils.config.remove_option("Favorites", option)

                for i, favorite in enumerate(favorites):
                    Utils.setStr("Favorites", f"favorite.{i}", favorite)

        Button(
            button_frame,
            text=_("Add"),
            command=add_favorite
        ).pack(side=LEFT, padx=5)

        Button(
            button_frame,
            text=_("Remove"),
            command=remove_favorite
        ).pack(side=LEFT, padx=5)

        Button(
            button_frame,
            text=_("Move Up"),
            command=move_up
        ).pack(side=LEFT, padx=5)

        Button(
            button_frame,
            text=_("Move Down"),
            command=move_down
        ).pack(side=LEFT, padx=5)

        Button(
            button_frame,
            text=_("Close"),
            command=dialog.destroy
        ).pack(side=RIGHT, padx=5)

        # Center dialog
        dialog.update_idletasks()
        w = dialog.winfo_width()
        h = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (w // 2)
        y = (dialog.winfo_screenheight() // 2) - (h // 2)
        dialog.geometry('{}x{}+{}+{}'.format(w, h, x, y))


# =============================================================================
# Options Group
# =============================================================================
class OptionsGroup(CNCRibbon.ButtonGroup):
    def __init__(self, master, app):
        CNCRibbon.ButtonGroup.__init__(self, master, N_("Options"), app)
        self.grid3rows()

        # ===
        col, row = 1, 0
        b = Ribbon.LabelButton(
            self.frame,
            text=_("Report"),
            image=Utils.icons["debug"],
            compound=LEFT,
            command=Utils.ReportDialog.sendErrorReport,
            anchor=W,
            background=Ribbon._BACKGROUND,
        )
        b.grid(row=row, column=col, padx=0, pady=0, sticky=EW)
        tkExtra.Balloon.set(b, _("Send Error Report"))

        # ---
        col, row = 1, 1
        b = Ribbon.LabelButton(
            self.frame,
            text=_("Updates"),
            image=Utils.icons["global"],
            compound=LEFT,
            command=self.app.checkUpdates,
            anchor=W,
            background=Ribbon._BACKGROUND,
        )
        b.grid(row=row, column=col, padx=0, pady=0, sticky=EW)
        tkExtra.Balloon.set(b, _("Check Updates"))

        col, row = 1, 2
        b = Ribbon.LabelButton(
            self.frame,
            text=_("About"),
            image=Utils.icons["about"],
            compound=LEFT,
            command=self.app.about,
            anchor=W,
            background=Ribbon._BACKGROUND,
        )
        b.grid(row=row, column=col, padx=0, pady=0, sticky=EW)
        tkExtra.Balloon.set(b, _("About the program"))


# =============================================================================
# Pendant Group
# =============================================================================
class PendantGroup(CNCRibbon.ButtonGroup):
    def __init__(self, master, app):
        CNCRibbon.ButtonGroup.__init__(self, master, N_("Pendant"), app)
        self.grid3rows()

        col, row = 0, 0
        b = Ribbon.LabelButton(
            self.frame,
            text=_("Start"),
            image=Utils.icons["start_pendant"],
            compound=LEFT,
            anchor=W,
            command=app.startPendant,
            background=Ribbon._BACKGROUND,
        )
        b.grid(row=row, column=col, padx=0, pady=0, sticky=NSEW)
        tkExtra.Balloon.set(b, _("Start pendant"))

        row += 1
        b = Ribbon.LabelButton(
            self.frame,
            text=_("Stop"),
            image=Utils.icons["stop_pendant"],
            compound=LEFT,
            anchor=W,
            command=app.stopPendant,
            background=Ribbon._BACKGROUND,
        )
        b.grid(row=row, column=col, padx=0, pady=0, sticky=NSEW)
        tkExtra.Balloon.set(b, _("Stop pendant"))


# =============================================================================
# Close Group
# =============================================================================
class CloseGroup(CNCRibbon.ButtonGroup):
    def __init__(self, master, app):
        CNCRibbon.ButtonGroup.__init__(self, master, N_("Close"), app)

        # ---
        b = Ribbon.LabelButton(
            self.frame,
            text=_("Exit"),
            image=Utils.icons["exit32"],
            compound=TOP,
            command=app.quit,
            anchor=W,
            background=Ribbon._BACKGROUND,
        )
        b.pack(fill=BOTH, expand=YES)
        tkExtra.Balloon.set(b, _("Close program [Ctrl-Q]"))


# =============================================================================
# Serial Frame
# =============================================================================
class SerialFrame(CNCRibbon.PageLabelFrame):
    def __init__(self, master, app):
        CNCRibbon.PageLabelFrame.__init__(
            self, master, "Serial", _("Serial"), app)
        self.autostart = BooleanVar()

        # ---
        col, row = 0, 0
        b = Label(self, text=_("Port:"))
        b.grid(row=row, column=col, sticky=E)
        self.addWidget(b)

        self.portCombo = tkExtra.Combobox(
            self,
            False,
            background=tkExtra.GLOBAL_CONTROL_BACKGROUND,
            width=16,
            command=self.comportClean,
        )
        self.portCombo.grid(row=row, column=col + 1, sticky=EW)
        tkExtra.Balloon.set(
            self.portCombo, _("Select (or manual enter) port to connect")
        )
        self.portCombo.set(Utils.getStr("Connection", "port"))
        self.addWidget(self.portCombo)

        self.comportRefresh()

        # ---
        row += 1
        b = Label(self, text=_("Baud:"))
        b.grid(row=row, column=col, sticky=E)

        self.baudCombo = tkExtra.Combobox(
            self, True, background=tkExtra.GLOBAL_CONTROL_BACKGROUND
        )
        self.baudCombo.grid(row=row, column=col + 1, sticky=EW)
        tkExtra.Balloon.set(self.baudCombo, _("Select connection baud rate"))
        self.baudCombo.fill(BAUDS)
        self.baudCombo.set(Utils.getStr("Connection", "baud", "115200"))
        self.addWidget(self.baudCombo)

        # ---
        row += 1
        b = Label(self, text=_("Controller:"))
        b.grid(row=row, column=col, sticky=E)

        self.ctrlCombo = tkExtra.Combobox(
            self,
            True,
            background=tkExtra.GLOBAL_CONTROL_BACKGROUND,
            command=self.ctrlChange,
        )
        self.ctrlCombo.grid(row=row, column=col + 1, sticky=EW)
        tkExtra.Balloon.set(self.ctrlCombo, _("Select controller board"))
        self.ctrlCombo.fill(self.app.controllerList())
        self.ctrlCombo.set(app.controller)
        self.addWidget(self.ctrlCombo)

        # ---
        row += 1
        b = Checkbutton(self, text=_("Connect on startup"),
                        variable=self.autostart)
        b.grid(row=row, column=col, columnspan=2, sticky=W)
        tkExtra.Balloon.set(
            b, _("Connect to serial on startup of the program"))
        self.autostart.set(Utils.getBool("Connection", "openserial"))
        self.addWidget(b)

        # ---
        col += 2
        self.comrefBtn = Ribbon.LabelButton(
            self,
            image=Utils.icons["refresh"],
            text=_("Refresh"),
            compound=TOP,
            command=lambda s=self: s.comportRefresh(True),
            background=Ribbon._BACKGROUND,
        )
        self.comrefBtn.grid(row=row, column=col, padx=0, pady=0, sticky=NSEW)
        tkExtra.Balloon.set(self.comrefBtn, _("Refresh list of serial ports"))

        # ---
        row = 0

        self.connectBtn = Ribbon.LabelButton(
            self,
            image=Utils.icons["serial48"],
            text=_("Open"),
            compound=TOP,
            command=lambda s=self: s.event_generate("<<Connect>>"),
            background=Ribbon._BACKGROUND,
        )
        self.connectBtn.grid(
            row=row, column=col, rowspan=3, padx=0, pady=0, sticky=NSEW
        )
        tkExtra.Balloon.set(self.connectBtn, _("Open/Close serial port"))
        self.grid_columnconfigure(1, weight=1)

    # -----------------------------------------------------------------------
    def ctrlChange(self):
        self.app.controllerSet(self.ctrlCombo.get())

    # -----------------------------------------------------------------------
    def comportClean(self, event=None):
        clean = self.portCombo.get().split("\t")[0]
        if self.portCombo.get() != clean:
            print("comport fix")
            self.portCombo.set(clean)

    # -----------------------------------------------------------------------
    def comportsGet(self):
        try:
            return comports(include_links=True)
        except TypeError:
            print("Using old style comports()!")
            return comports()

    def comportRefresh(self, dbg=False):
        # Detect devices
        hwgrep = []
        for i in self.comportsGet():
            if dbg:
                # Print list to console if requested
                comport = ""
                for j in i:
                    comport += j + "\t"
                print(comport)
            for hw in i[2].split(" "):
                hwgrep += ["hwgrep://" + hw + "\t" + i[1]]

        # Populate combobox
        devices = sorted(x[0] + "\t" + x[1] for x in self.comportsGet())
        devices += [""]
        devices += sorted(set(hwgrep))
        devices += [""]
        # Pyserial raw spy currently broken in python3
        # TODO: search for python3 replacement for raw spy
        if sys.version_info[0] != 3:
            devices += sorted(
                "spy://" + x[0] + "?raw&color" + "\t(Debug) " + x[1]
                for x in self.comportsGet()
            )
        else:
            devices += sorted(
                "spy://" + x[0] + "?color" + "\t(Debug) " + x[1]
                for x in self.comportsGet()
            )
        devices += ["", "socket://localhost:23", "rfc2217://localhost:2217"]

        # Clean neighbour duplicates
        devices_clean = []
        devprev = ""
        for i in devices:
            if i.split("\t")[0] != devprev:
                devices_clean += [i]
            devprev = i.split("\t")[0]

        self.portCombo.fill(devices_clean)

    # -----------------------------------------------------------------------
    def saveConfig(self):
        # Connection
        Utils.setStr("Connection", "controller", self.app.controller)
        Utils.setStr("Connection", "port", self.portCombo.get().split("\t")[0])
        Utils.setStr("Connection", "baud", self.baudCombo.get())
        Utils.setBool("Connection", "openserial", self.autostart.get())


# =============================================================================
# File Page
# =============================================================================
class FilePage(CNCRibbon.Page):
    __doc__ = _("File I/O and configuration")
    _name_ = N_("File")
    _icon_ = "new"

    # ----------------------------------------------------------------------
    # Add a widget in the widgets list to enable disable during the run
    # ----------------------------------------------------------------------
    def register(self):
        self._register(
            (FileGroup, PendantGroup, OptionsGroup, CloseGroup), (SerialFrame,)
        )