using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

// Reads and drives a SysTreeView32 that lives in another process (Excel).
public static class TreeRemote {
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("kernel32.dll")] static extern IntPtr OpenProcess(uint a, bool i, uint pid);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
    [DllImport("kernel32.dll")] static extern IntPtr VirtualAllocEx(IntPtr p, IntPtr a, UIntPtr s, uint t, uint pr);
    [DllImport("kernel32.dll")] static extern bool VirtualFreeEx(IntPtr p, IntPtr a, UIntPtr s, uint t);
    [DllImport("kernel32.dll")] static extern bool WriteProcessMemory(IntPtr p, IntPtr a, byte[] b, UIntPtr s, out UIntPtr w);
    [DllImport("kernel32.dll")] static extern bool ReadProcessMemory(IntPtr p, IntPtr a, byte[] b, UIntPtr s, out UIntPtr r);

    const uint TVM_EXPAND = 0x1102, TVM_GETITEMRECT = 0x1104, TVM_GETNEXTITEM = 0x110A,
               TVM_SELECTITEM = 0x110B, TVM_ENSUREVISIBLE = 0x1114, TVM_GETITEMW = 0x113E;
    const int TVGN_ROOT = 0, TVGN_NEXT = 1, TVGN_CHILD = 4, TVGN_CARET = 9;

    static IntPtr Next(IntPtr tree, IntPtr item, int code) {
        return SendMessage(tree, TVM_GETNEXTITEM, (IntPtr)code, item);
    }

    public static string Text(IntPtr tree, IntPtr item) {
        uint pid; GetWindowThreadProcessId(tree, out pid);
        IntPtr proc = OpenProcess(0x0008 | 0x0010 | 0x0020, false, pid); // VM_OPERATION|VM_READ|VM_WRITE
        int structSize = 56, textChars = 260, total = structSize + textChars * 2;
        IntPtr mem = VirtualAllocEx(proc, IntPtr.Zero, (UIntPtr)total, 0x1000 | 0x2000, 0x04);
        try {
            byte[] s = new byte[structSize];
            BitConverter.GetBytes((uint)1).CopyTo(s, 0);                      // mask = TVIF_TEXT
            BitConverter.GetBytes(item.ToInt64()).CopyTo(s, 8);               // hItem
            BitConverter.GetBytes(mem.ToInt64() + structSize).CopyTo(s, 24); // pszText
            BitConverter.GetBytes(textChars).CopyTo(s, 32);                  // cchTextMax
            UIntPtr n; WriteProcessMemory(proc, mem, s, (UIntPtr)structSize, out n);
            SendMessage(tree, TVM_GETITEMW, IntPtr.Zero, mem);
            byte[] t = new byte[textChars * 2];
            ReadProcessMemory(proc, mem + structSize, t, (UIntPtr)t.Length, out n);
            string str = Encoding.Unicode.GetString(t);
            int z = str.IndexOf('\0'); return z >= 0 ? str.Substring(0, z) : str;
        } finally {
            VirtualFreeEx(proc, mem, UIntPtr.Zero, 0x8000); CloseHandle(proc);
        }
    }

    public static List<string> Dump(IntPtr tree, int maxDepth) {
        var res = new List<string>();
        Walk(tree, Next(tree, IntPtr.Zero, TVGN_ROOT), 0, maxDepth, res);
        return res;
    }

    static void Walk(IntPtr tree, IntPtr item, int depth, int maxDepth, List<string> res) {
        while (item != IntPtr.Zero && res.Count < 500) {
            res.Add(new string(' ', depth * 2) + Text(tree, item));
            if (depth < maxDepth) Walk(tree, Next(tree, item, TVGN_CHILD), depth + 1, maxDepth, res);
            item = Next(tree, item, TVGN_NEXT);
        }
    }

    // Finds a node by a path of alternatives, e.g. { {"MyApp"}, {"Dimensions","Dimensiones"}, {"Account"} },
    // expanding each level. Returns IntPtr.Zero when a level is not found.
    public static IntPtr Find(IntPtr tree, string[][] path) {
        IntPtr level = Next(tree, IntPtr.Zero, TVGN_ROOT);
        IntPtr found = IntPtr.Zero;
        for (int p = 0; p < path.Length; p++) {
            found = SearchDown(tree, level, path[p], 0);
            if (found == IntPtr.Zero) return IntPtr.Zero;
            SendMessage(tree, TVM_EXPAND, (IntPtr)2, found);
            System.Threading.Thread.Sleep(400);
            level = Next(tree, found, TVGN_CHILD);
        }
        return found;
    }

    // Breadth: siblings at this level; depth: the first path step may sit under the server node.
    static IntPtr SearchDown(IntPtr tree, IntPtr first, string[] names, int depth) {
        for (IntPtr it = first; it != IntPtr.Zero; it = Next(tree, it, TVGN_NEXT)) {
            string t = Text(tree, it);
            foreach (var n in names) if (string.Equals(t, n, StringComparison.OrdinalIgnoreCase)) return it;
        }
        if (depth >= 2) return IntPtr.Zero;
        for (IntPtr it = first; it != IntPtr.Zero; it = Next(tree, it, TVGN_NEXT)) {
            IntPtr c = Next(tree, it, TVGN_CHILD);
            if (c == IntPtr.Zero) continue;
            IntPtr r = SearchDown(tree, c, names, depth + 1);
            if (r != IntPtr.Zero) return r;
        }
        return IntPtr.Zero;
    }

    // Selects the node and presses Enter on it: this is what opens a node in the Smart
    // View panel (synthetic mouse clicks are ignored).
    public static void Open(IntPtr tree, IntPtr item) {
        SendMessage(tree, TVM_SELECTITEM, (IntPtr)TVGN_CARET, item);
        SendMessage(tree, TVM_ENSUREVISIBLE, IntPtr.Zero, item);
        PostMessage(tree, 0x100, (IntPtr)0x0D, (IntPtr)0x001C0001); // WM_KEYDOWN VK_RETURN
        PostMessage(tree, 0x101, (IntPtr)0x0D, unchecked((IntPtr)(long)0xC01C0001)); // WM_KEYUP
    }

    // Finds the visible SysTreeView32 of the Smart View panel in an Excel process.
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc p, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc p, IntPtr l);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    delegate bool EnumProc(IntPtr h, IntPtr l);
    public static IntPtr FindPanelTree(uint excelPid) {
        IntPtr found = IntPtr.Zero;
        EnumWindows((h, l) => {
            uint p; GetWindowThreadProcessId(h, out p);
            if (p != excelPid) return true;
            EnumChildWindows(h, (c, l2) => {
                var sb = new StringBuilder(64); GetClassName(c, sb, 64);
                if (sb.ToString() == "SysTreeView32" && IsWindowVisible(c) && Next(c, IntPtr.Zero, TVGN_ROOT) != IntPtr.Zero) { found = c; return false; }
                return true;
            }, IntPtr.Zero);
            return found == IntPtr.Zero;
        }, IntPtr.Zero);
        return found;
    }
}
