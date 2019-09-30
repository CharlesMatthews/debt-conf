#!/usr/bin/python3


import os
from pylatex import Document, Head, Tabular, MiniPage, LargeText, LineBreak, MediumText
from pylatex.utils import italic, bold, NoEscape

def make_pdf(auditor_name, auditor_email, address_list, invoices, outfile="outfile", letter_head=None, signature_file=None, client_signatory=None):
    print(letter_head)
    print(signature_file)
    cust_name = address_list[0]
    tex_address = r"\\".join(address_list)

    ###### BEGIN HARDCODED LATEX DOCUMENT !
    preamble = r"""
    \documentclass{letter}


    % preamble
    \usepackage{geometry}
    \geometry{%
    a4paper,
    left=20 mm,
    right=20 mm,
    top=20 mm,
    bottom=20 mm
    }

    %\usepackage[a4paper, margin=1in]{geometry}


    % allow graphics and pdfs
    \usepackage{graphicx}
    \usepackage{pdfpages}

    % allow colours
    %\usepackage[dvipsnames]{xcolor}

    % allow lists
    \usepackage{enumitem}

    % allow hyperlinks
    \usepackage{hyperref}
    \usepackage{fancyref}


    % Helvet font (uncomment if required)
    %\usepackage{sansmathfonts}
    \usepackage{helvet}
    \renewcommand{\familydefault}{\sfdefault}
    \usepackage[T1]{fontenc}
    \usepackage{textcomp}

    \usepackage{array}


    \makeatletter
    \def\opening#1{\ifx\@empty\fromaddress
      \thispagestyle{firstpage}%
        {\raggedleft\@date\par}%
      \else  % home address
        \thispagestyle{empty}%
        {\noindent\let\\\cr\halign{##\hfil\cr\ignorespaces
          \fromaddress \cr\noalign{\kern 2\parskip}%
          \@date\cr}\par}%
      \fi
      \vspace{2\parskip}%
      {\raggedright \toname \\ \toaddress \par}%
      \vspace{2\parskip}%
      #1\par\nobreak}
    \makeatother
    """

    if letter_head is not None:
        preamble += r"""
        \usepackage{eso-pic}

        \newcommand\BackgroundPic{
            \put(0,0){
            \parbox[b][\paperheight]{\paperwidth}{%
            \vfill
            \centering
            \includegraphics[width=\paperwidth,height=\paperheight]{""" + letter_head + r"""}
            \vfill
            }}}"""

    preamble+= r"""\signature{"""

    if signature_file is not None:
        preamble += r"""\vspace{-2cm}"""

    preamble += client_signatory + r""" }
    \address{""" + tex_address + r"""}
    \longindentation=0pt
    \begin{document}"""

    if letter_head is not None:
        preamble+= r"""
        \AddToShipoutPicture*{\BackgroundPic}
            \setlength{\extrarowheight}{10pt}"""

    preamble += r"""
    \begin{letter}{\bf{RE: Unpaid Invoice Follow-up}}
    \opening{To Whom it may concern,}

    Generic text here to explain problem.

        You can find """ + auditor_name + r""" at \href{mailto:""" + auditor_email + r"}{" + auditor_email + r"""}.


    \begin{center}
    \begin{tabular}{p{2cm} p{4cm} p{4cm}}
        & \bf{Invoice Number} & \bf{Amount} \\[0.5ex]
       \hline
    """

    datatex = r""
    for i, inv in enumerate(invoices):
        datatex += str(i+1) + " & " + str(inv["number"]) + "&" + str(inv["currency"]) + " " + str(inv["value"]) + r"\\  "

#demo_entry =  r"1 & 60123 & GPB 787 \\"

    bottom = r"""
    \end{tabular}
    \end{center}

        \closing{Warm Regards, """
    if signature_file is not None:
        bottom += r"""\\ \includegraphics[width=4cm,height=3cm,keepaspectratio]{""" + signature_file + r"""}"""

    bottom += r"""}
    \end{letter}
    \end{document}
    """
    ###### END HARDCODED LATEX DOCUMENT !

    tex = preamble + datatex + bottom

    orig_dir = os.getcwd()

    os.chdir(r"""Output\PDFs""")
    print(("IN " + os.getcwd()+"\n")*10)

    real_outfile = outfile+".tex"
    with open(real_outfile,"w") as f:
        f.write(tex)

        print((real_outfile+"\n")*10)

    os.system("pdflatex " + real_outfile) ### THIS WILL CHANGE WHEN USING WINDOWS !!!!
    os.chdir(orig_dir)
    print(("OUT " + os.getcwd()+"\n")*10)



if __name__ == "__main__":
    address_list = ["CustomerName", "123 demo st", "London", "SW1 000", "UK"]
    invoices = [{"number":123, "value":1231.12, "currency":"GBP"}, {"number":456, "value":98765.12, "currency":"USD"}]
    auditor_name = "Charles Matthews"
    auditor_email = "charles.matthews.test@ey.com"
    make_pdf(auditor_name, auditor_email, address_list, invoices, outfile="demo", letter_head = "lhead.pdf", signature_file="signature.png")
