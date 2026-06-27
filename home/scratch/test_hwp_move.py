import win32com.client as win32
import os

source_path = os.path.abspath(r'C:\Users\-\OneDrive\바탕 화면\1. 안전보건협의체 수급업체 회의자료(서식).hwp')

try:
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    hwp.Open(source_path)
    
    hwp.HAction.Run("MoveDocBegin")
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = "수급업체명"
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("Forward")
    hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    
    hwp.HAction.Run("MoveRight") # deselect
    
    # Get current cell text
    hwp.HAction.Run("SelectAll")
    hwp.HAction.Run("Copy")
    # Actually, getting text via API is better
    # But just let's see if TableRightCell works
    
    hwp.HAction.Run("Cancel")
    print("Executing TableRightCell...")
    hwp.HAction.Run("TableRightCell")
    
    hwp.HAction.Run("SelectAll")
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = "TEST_SUCCESS"
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    
    output_path = os.path.abspath(r'C:\Users\-\OneDrive\바탕 화면\test_hwp_out.hwp')
    hwp.SaveAs(output_path)
    print("Done")
except Exception as e:
    print(f"Error: {e}")
finally:
    try:
        hwp.Quit()
    except:
        pass
