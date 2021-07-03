$(function() {
    $('.menu-icon').click(function(e) {
        $('.wrapper-nav').toggleClass('wrapper-nav2');
        e.preventDefault();
    });
});

function myFunction() {

    var replacement = "<span class='blue'>repla</span>";

    var tablica = ["True", "False", "Sub", "Public Sub", "Public", "Private Sub", "End ", "If", "End Sub", "End If", "As ", "Dim ", "Nothing", "Is ", "New", "Empty", "Set ", "Function", "String", "Integer", "On Error GoTo", "On Error GoTo 0", " On Error Resume Next", "For ", "For Each", "In ", "Next", "Exit ", "End", "Or ", "And ", "Then", "Else ", "Long ", "To ", "Do ", "Until", "Loop", "Wend", "While", "With", " Select", " Case", "Else", "Else If", "Boolean", "Variant", "Resume", "Call", "Object ", "Optional", "If"];

    for (i=0;i<tablica.length;i++){

        var str = document.getElementById("demo").innerHTML;

        var replacethis = tablica[i];

        var re = new RegExp(replacethis,"g");

        var reeee = new RegExp(replacement,"g");

        var txt = str.replace(re, replacement.replace("repla",replacethis) );

        document.getElementById("demo").innerHTML = txt;

    }

}

myFunction();

