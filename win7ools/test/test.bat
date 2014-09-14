
pandoc -s -S resume.txt -o resume.docx

pandoc -s resume.txt --template "C:\xampp\htdocs\resume.jasonwohlgemuth\formats\templates\resume.template.html.html" -t html5 -o resume.html

pandoc -s resume.txt --template "C:\xampp\htdocs\resume.jasonwohlgemuth\formats\templates\resume.template.tex.tex" -o resume.tex

pandoc resume.txt --template "C:\xampp\htdocs\resume.jasonwohlgemuth\formats\templates\resume.template.tex.tex" -o resume.pdf
