import tkinter


class EntryBox(object):

    root = None

    def __init__(self, msg, dict_key=None):
        """
        msg = <str> the message to be displayed
        dict_key = <sequence> (dictionary, key) to associate with user input
        (providing a sequence for dict_key creates an entry for user input)
        """
        tki = tkinter
        self.top = tki.Toplevel(EntryBox.root)

        frm = tki.Frame(self.top)
        frm.pack(fill='both', expand=True)

        label = tki.Label(frm, text=msg)
        label.pack(padx=8, pady=8)

        caller_wants_an_entry = dict_key is not None

        if caller_wants_an_entry:
            self.entry = tki.Entry(frm)
            d, key = dict_key
            self.entry.insert(0, d[key])
            self.entry.focus_set()
            self.entry.pack(fill='x', padx=8, pady=8)
            self.entry.bind('<Return>', lambda x: self.entry_to_dict(dict_key))

            b_submit = tki.Button(frm, text='OK', width=10)
            b_submit['command'] = lambda: self.entry_to_dict(dict_key)
            b_submit.pack(side='left', padx=8, pady=8)

            b_cancel = tki.Button(frm, text='Cancel', width=10)
            b_cancel['command'] = self.top.destroy
            b_cancel.pack(side='left', padx=8, pady=8)

        self.centre()
        self.top.grab_set()

    def entry_to_dict(self, dict_key):
        data = self.entry.get()
        if data:
            d, key = dict_key
            d[key] = data
            self.top.destroy()

    def centre(self):
        """Center the window on screen."""
        self.top.update_idletasks()
        # The horizontal position is calculated as (screenwidth - window_width)/2
        hpos = int((self.top.winfo_screenwidth() - self.top.winfo_width())/2)
        # And vertical position the same, but with the height dimensions
        vpos = int((self.top.winfo_screenheight() - self.top.winfo_height())/2)
        # And the move call repositions the window
        self.top.geometry('+{x}+{y}'.format(x=hpos, y=vpos))
