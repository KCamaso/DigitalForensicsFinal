Imports System.IO
Imports System.IO.Compression
Imports System.Xml


Module Program
    Sub Main(args As String())
        Console.WriteLine("DOCX Edit Viewer")
        Console.WriteLine("by Karen Armas Camaso")

        Console.WriteLine("Please input the file path of the docx followed by .docx (Example: 'c:\Documents\document.docx')")
        Dim zipPath As String = Console.ReadLine()

        Console.WriteLine("Please input the path of where the extracted files will go (Example: 'c:\Documents\Output')")
        Dim extractPath As String = Console.ReadLine()

        ZipFile.ExtractToDirectory(zipPath, extractPath, True)

        Dim xr As XmlReader = XmlReader.Create(extractPath + "\docProps\core.xml")
        Dim xa As XmlReader = XmlReader.Create(extractPath + "\word\document.xml")

        End If
        Dim idEdit As New List(Of String)
        While xa.Read()
            If xa.Name = "w:r" Then
                Dim id As String = xa("w:rsidR")
                If id IsNot Nothing Then
                    idEdit.Add(id)

                End If
            End If
            If xa.Name = "w:p" Then
                Dim id As String = xa("w:rsidR")
                If id IsNot Nothing Then
                    idEdit.Add(id)

                End If
            End If
        End While

        Dim group = idEdit.GroupBy(Function(value) value)


        Console.WriteLine("")

        Console.WriteLine("Edit IDs (In Order of Appearance)")
        For Each grp In group
            Console.WriteLine(grp(0) & " - " & grp.Count & " times")
        Next

        While xr.Read()

            If xr.IsStartElement("dc:creator") Then
                Console.WriteLine("Created by: {0}", xr.ReadInnerXml())
            End If
            If xr.IsStartElement("dc:lastModifiedBy") Then
                Console.WriteLine("Last Modified by: {0}", xr.ReadInnerXml())
            End If
            If xr.IsStartElement("dcterms:created") Then
                Console.WriteLine("Created on: {0}", xr.ReadInnerXml())
            End If
            If xr.IsStartElement("dcterms:modified") Then
                Console.WriteLine("Last Modified on: {0}", xr.ReadInnerXml())
            End If
        End While



        Console.WriteLine("Press any key to continue . . . ")

        Console.ReadKey(True)
    End Sub

End Module
