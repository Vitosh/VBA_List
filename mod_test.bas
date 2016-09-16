Attribute VB_Name = "mod_test"
Option Explicit

Sub TestClassVbaList()

    Dim obj_list As cls_vbaList

    Set obj_list = New cls_vbaList

    obj_list.Add (30)
    obj_list.Add (3)
    obj_list.Add (355)
    obj_list.Add (5)
    obj_list.Add (1)
    obj_list.Add (40)

    Debug.Print obj_list.Contains(30)
    Debug.Print obj_list.Exists(30)
    Debug.Print obj_list.Items(0)
    
    obj_list.Sort
    
    Debug.Print obj_list.Items(0)
    Debug.Print obj_list.Find(3)
    Debug.Print obj_list.Find(30)
    Debug.Print obj_list.LastIndexOf(355)

    Set obj_list = Nothing

End Sub
