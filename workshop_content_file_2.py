"""
Content File for Workshop
"""
def base_content(s_date, e_date, gender, collabrator):
        
        if gender == "M":
            text = ""
        elif gender == "F":
           text = ""

   
        return text


def text_wrap(text, font, max_width):
    lines = []
    # If the width of the text is smaller than image width
    # we don't need to split it, just add it to the lines array
    # and return
    if font.getsize(text)[0] <= max_width:
        lines.append(text)
    else:
        # split the line by spaces to get words
        words = text.split(' ')
        i = 0
        # append every word to a line while its width is shorter than image width
        while i < len(words):
            line = ''
            while i < len(words) and font.getsize(line + words[i])[0] <= max_width:
                line = line + words[i] + " "
                i += 1
            if not line:
                line = words[i]
                i += 1
            # when the line gets longer than the max width do not append the word,
            # add the line to the lines array
            lines.append(line)
    return lines

#import textwrap

def certificate_content(name, gender, college, font, page_width, start_date, end_date, title, collabrator):
    content = base_content(start_date, end_date, gender, collabrator).format(name, title, collabrator, start_date,end_date,college).split("\n")
    final_content = []
    for para in content:
#        final_text.append(textwrap.wrap(i,85))
        final_content.append(text_wrap(para, font, page_width))
#    print final_text
    return final_content
