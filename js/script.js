$('a[href*=\\#]').on('click', function(event) {     
    event.preventDefault();

    $('html,body').animate({scrollTop:$(this.hash).offset().top}, 900);
});
var nameValue = document.getElementById("fback").value;
function WriteToFile(passForm) {
    set fso = CreateObject("Scripting.FileSystemObject");  
    set s = fso.CreateTextFile("C:\test.txt", True);
    s.writeline("HI");
    s.writeline("Bye");
    s.writeline("-----------------------------");
    s.Close();
 }
