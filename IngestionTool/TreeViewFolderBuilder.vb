Module TreeViewFolderBuilder

    Private _ImageKey As System.String
    Private _ExpandedImageKey As System.String
    Private _ContextMenuStrip As System.Windows.Forms.ContextMenuStrip

    Public Sub Initialize(TreeView As System.Windows.Forms.TreeView, RootDirectory As System.String)
        Initialize(TreeView, New System.IO.DirectoryInfo(RootDirectory), System.String.Empty, System.String.Empty, Nothing)
    End Sub
    Public Sub Initialize(TreeView As System.Windows.Forms.TreeView, RootDirectory As System.String, ImageKey As System.String, ExpandedImageKey As System.String)
        Initialize(TreeView, New System.IO.DirectoryInfo(RootDirectory), ImageKey, ExpandedImageKey, Nothing)
    End Sub
    Public Sub Initialize(TreeView As System.Windows.Forms.TreeView, RootDirectory As System.String, ImageKey As System.String, ExpandedImageKey As System.String, ContextMenuStrip As System.Windows.Forms.ContextMenuStrip)
        Initialize(TreeView, New System.IO.DirectoryInfo(RootDirectory), ImageKey, ExpandedImageKey, ContextMenuStrip)
    End Sub
    Public Sub Initialize(TreeView As System.Windows.Forms.TreeView, RootDirectoryInfo As System.IO.DirectoryInfo)
        Initialize(TreeView, RootDirectoryInfo, System.String.Empty, System.String.Empty, Nothing)
    End Sub
    Public Sub Initialize(TreeView As System.Windows.Forms.TreeView, RootDirectoryInfo As System.IO.DirectoryInfo, ImageKey As System.String, ExpandedImageKey As System.String, ContextMenuStrip As System.Windows.Forms.ContextMenuStrip)
        _ImageKey = ImageKey
        _ExpandedImageKey = ExpandedImageKey
        _ContextMenuStrip = ContextMenuStrip
        AddHandler TreeView.AfterCollapse, AddressOf TreeView_AfterCollapse
        AddHandler TreeView.AfterExpand, AddressOf TreeView_AfterExpand
        TreeView.Nodes.Clear()
        TreeView.BeginUpdate()
        TreeView.Nodes.Add(CreateTreeNode(RootDirectoryInfo, True))
        TreeView.EndUpdate()
    End Sub

    Private Function CreateTreeNode(DirectoryInfo As System.IO.DirectoryInfo, Expanded As System.Boolean) As System.Windows.Forms.TreeNode
        Dim TreeNode As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode
        TreeNode.Tag = DirectoryInfo
        TreeNode.ContextMenuStrip = _ContextMenuStrip
        TreeNode.Text = DirectoryInfo.Name
        If Expanded = True Then
            TreeNode.ImageKey = _ExpandedImageKey
            For Each ChildDirectoryInfo As System.IO.DirectoryInfo In DirectoryInfo.GetDirectories("*", IO.SearchOption.TopDirectoryOnly)
                TreeNode.Nodes.Add(CreateTreeNode(ChildDirectoryInfo, False))
            Next
            TreeNode.Expand()
        Else
            TreeNode.ImageKey = _ImageKey
            TreeNode.Nodes.Add(CreateLoadingTreeNode)
        End If
        Return TreeNode
    End Function

    Private Function CreateLoadingTreeNode()
        Dim TreeNode As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode
        TreeNode.Text = "Loading..."
        Return TreeNode
    End Function

    Private Sub TreeView_AfterCollapse(sender As Object, e As TreeViewEventArgs)
        e.Node.ImageKey = _ImageKey
    End Sub

    Private Sub TreeView_AfterExpand(sender As Object, e As TreeViewEventArgs)
        e.Node.ImageKey = _ExpandedImageKey
        If e.Node.Nodes.Count = 1 AndAlso e.Node.Nodes.Item(0).Tag Is Nothing AndAlso e.Node.Nodes.Item(0).Text = "Loading..." Then
            Application.DoEvents()
            e.Node.TreeView.BeginUpdate()
            e.Node.Nodes.Clear()
            Dim DirectoryInfo As System.IO.DirectoryInfo = DirectCast(e.Node.Tag, System.IO.DirectoryInfo)
            For Each ChildDirectoryInfo As System.IO.DirectoryInfo In DirectoryInfo.GetDirectories("*", IO.SearchOption.TopDirectoryOnly)
                e.Node.Nodes.Add(CreateTreeNode(ChildDirectoryInfo, False))
            Next
            e.Node.TreeView.EndUpdate()
        End If
    End Sub

End Module
