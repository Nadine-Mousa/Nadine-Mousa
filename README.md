### Hi there 👋

I'm Nadine, a passionate software engineer with a love for building amazing software and solving complex problems. 

![](https://komarev.com/ghpvc/?username=Nadine-Mousa)

<!--
**Nadine-Mousa/Nadine-Mousa** is a ✨ _special_ ✨ repository because its `README.md` (this file) appears on your GitHub profile.


Here are some ideas to get you started:

- 🔭 I’m currently working on ...
- 🌱 I’m currently learning ...
- 👯 I’m looking to collaborate on ...
- 🤔 I’m looking for help with ...
- 💬 Ask me about ...
- 📫 How to reach me: ...
- 😄 Pronouns: ...
- ⚡ Fun fact: ...





Sub CreatePresentation()
    Dim pptApp As Object
    Dim pptPres As Object
    Dim slideIndex As Integer

    ' Create a new instance of PowerPoint application
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True

    ' Create a new presentation
    Set pptPres = pptApp.Presentations.Add

    ' Add title slide
    slideIndex = slideIndex + 1
    AddTitleSlide pptPres.Slides.Add(slideIndex, ppLayoutTitle), "Data Structures and Algorithms in C#", "Your Name", "University Name", "Date"

    ' Add content slides
    slideIndex = slideIndex + 1
    AddContentSlide pptPres.Slides.Add(slideIndex, ppLayoutText), "Introduction to Data Structures", "Data structures are essential in programming for organizing and storing data efficiently."

    slideIndex = slideIndex + 1
    AddContentSlide pptPres.Slides.Add(slideIndex, ppLayoutText), "Arrays", "Arrays are one of the most basic data structures in C#, providing a contiguous block of memory to store elements of the same type."

    slideIndex = slideIndex + 1
    AddContentSlide pptPres.Slides.Add(slideIndex, ppLayoutText), "Linked Lists", "Linked lists consist of nodes where each node contains a data field and a reference (link) to the next node in the sequence."

    slideIndex = slideIndex + 1
    AddContentSlide pptPres.Slides.Add(slideIndex, ppLayoutText), "Stacks", "A stack is a Last-In-First-Out (LIFO) data structure where elements are added and removed from the top."

    slideIndex = slideIndex + 1
    AddContentSlide pptPres.Slides.Add(slideIndex, ppLayoutText), "Queues", "A queue is a First-In-First-Out (FIFO) data structure where elements are added at the rear and removed from the front."

    slideIndex = slideIndex + 1
    AddContentSlide pptPres.Slides.Add(slideIndex, ppLayoutText), "Binary Trees", "Binary trees are hierarchical data structures consisting of nodes, each having at most two children, referred to as the left child and the right child."

    ' Add problem-solving slides
    slideIndex = slideIndex + 1
    AddContentSlide pptPres.Slides.Add(slideIndex, ppLayoutText), "Problem 1", "Implement a function to reverse a singly linked list."

    slideIndex = slideIndex + 1
    AddContentSlide pptPres.Slides.Add(slideIndex, ppLayoutText), "Solution 1", "Iterate through the linked list, reversing the links between nodes."

    slideIndex = slideIndex + 1
    AddContentSlide pptPres.Slides.Add(slideIndex, ppLayoutText), "Problem 2", "Implement a function to check if a given binary tree is a binary search tree."

    slideIndex = slideIndex + 1
    AddContentSlide pptPres.Slides.Add(slideIndex, ppLayoutText), "Solution 2", "Perform an inorder traversal of the binary tree while maintaining the previous node's value. If the current node's value is less than or equal to the previous node's value, the tree is not a binary search tree."

    ' Save the presentation
    pptPres.SaveAs "Data_Structures_Algorithms_CSharp.pptx"

    ' Clean up
    pptPres.Close
    pptApp.Quit
    Set pptPres = Nothing
    Set pptApp = Nothing
End Sub

Function AddTitleSlide(slide As Object, title As String, author As String, university As String, dateStr As String)
    With slide.Shapes(1).TextFrame.TextRange
        .Text = title
        .Font.Size = 32
        .Font.Bold = True
        .ParagraphFormat.Alignment = 2 'Center align
    End With
    With slide.Shapes(2).TextFrame.TextRange
        .Text = "Author: " & author & vbCrLf & "University: " & university & vbCrLf & "Date: " & dateStr
        .Font.Size = 14
        .ParagraphFormat.Alignment = 2 'Center align
    End With
End Function

Function AddContentSlide(slide As Object, title As String, content As String)
    With slide.Shapes(1).TextFrame.TextRange
        .Text = title
        .Font.Size = 24
        .Font.Bold = True
        .ParagraphFormat.Alignment = 2 'Center align
    End With
    With slide.Shapes(2).TextFrame.TextRange
        .Text = content
        .Font.Size = 18
        .ParagraphFormat.Alignment = 2 'Center align
    End With
End Function

-->
