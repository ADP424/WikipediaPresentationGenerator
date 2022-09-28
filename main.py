import codecs
import re
from uuid import uuid4

import wikipedia
from wikipedia import WikipediaPage

from slides import get_credentials, presentation

1
def wiki_search(term: str):
    term_list = wikipedia.search(term)

    if len(term_list) == 0:
        print("The term did not produce any search results.")
        return None

    if len(term_list) == 1:
        return wikipedia.page(term)

    for i in range(len(term_list)):
        print(str(i + 1) + ". " + term_list[i])

    chosen_index = int(input("Please enter the number of the page you want: "))

    while chosen_index < 0 or chosen_index > len(term_list):
        chosen_index = input("That number is out of range. Enter again.")

    return wikipedia.page(term_list[chosen_index - 1])


def paragraph_to_sentence_list(paragraph: str) -> list:
    """
    Most of the code is completely taken with minor adaptations from here:
    https://stackoverflow.com/questions/68676571/how-to-split-text-into-sentences-by-including-corner-cases
    :param paragraph: the paragraph to be split into sentences
    :return: the list of sentences
    """

    alphabets = "([A-Za-z])"
    prefixes = "(Mr|St|Mrs|Ms|DrProf|Capt|Cpt|Lt|Mt)[.]"
    suffixes = "(Inc|Ltd|Jr|Sr|Co)"
    starters = "(Mr|Mrs|Ms|Dr|He\s|She\s|It\s|They\s|Their\s|Our\s|We\s|But\s|However\s|That\s|This\s|Wherever)"
    acronyms = "([A-Z][.][A-Z][.](?:[A-Z][.])?)"
    websites = "[.](com|net|org|io|gov|me|edu)"
    digitPrefixes = "(No)"
    digits = "([0-9])"

    paragraph = " " + paragraph + "  "
    paragraph = paragraph.replace("\n", ".<stop>")
    paragraph = re.sub(prefixes, "\\1<prd>", paragraph)
    paragraph = re.sub(websites, "<prd>\\1", paragraph)
    if "Ph.D" in paragraph:
        paragraph = paragraph.replace("Ph.D.", "Ph<prd>D<prd>")
    paragraph = re.sub("\s" + alphabets + "[.] ", " \\1<prd> ", paragraph)
    paragraph = re.sub(acronyms + " " + starters, "\\1<stop> \\2", paragraph)
    paragraph = re.sub(alphabets + "[.]" + alphabets + "[.]" + alphabets + "[.]", "\\1<prd>\\2<prd>\\3<prd>", paragraph)
    paragraph = re.sub(alphabets + "[.]" + alphabets + "[.]", "\\1<prd>\\2<prd>", paragraph)
    paragraph = re.sub(" " + suffixes + "[.] " + starters, " \\1<stop> \\2", paragraph)
    paragraph = re.sub(" " + suffixes + "[.]", " \\1<prd>", paragraph)
    paragraph = re.sub(" " + alphabets + "[.]", " \\1<prd>", paragraph)
    paragraph = re.sub(digits + "[.]" + digits, "\\1<prd>\\2", paragraph)
    paragraph = re.sub(digitPrefixes + "[.] " + digits, "\\1<prd> \\2", paragraph)
    paragraph = re.sub(digitPrefixes + "[.] " + digits + digits, "\\1<prd> \\2\\3", paragraph)
    paragraph = re.sub(digitPrefixes + "[.] " + digits + digits + digits, "\\1<prd> \\2\\3\\4", paragraph)

    if "”" in paragraph:
        paragraph = paragraph.replace(".”", "”.")
    if "\"" in paragraph:
        paragraph = paragraph.replace(".\"", "\".")
    if "!" in paragraph:
        paragraph = paragraph.replace("!\"", "\"!")
    if "?" in paragraph:
        paragraph = paragraph.replace("?\"", "\"?")
    paragraph = paragraph.replace(".", ".<stop>")
    paragraph = paragraph.replace("?", "?<stop>")
    paragraph = paragraph.replace("!", "!<stop>")
    paragraph = paragraph.replace("<prd>", ".")
    sentences = paragraph.split("<stop>")

    i = 0
    while i < len(sentences):
        sentences[i] = sentences[i][:-1] # removes the period at the end
        if len(sentences[i]) < 10 or len(sentences[i]) > 200:
            sentences.remove(sentences[i])
            continue

        i += 1

    return sentences


def get_bullet_points(page: WikipediaPage, header: str, num: int) -> str:
    bullet_points = ""
    if header.lower() == "summary":
        sentences = paragraph_to_sentence_list(page.summary)
    else:
        sentences = paragraph_to_sentence_list(find_body_of_header(page, header))

    for i in range(num):
        if len(sentences) > i:
            bullet_points += sentences[i].strip() + "\n"

    return bullet_points.strip()


def write_article_to_text(page: WikipediaPage, file: codecs.StreamReaderWriter):
    file.write(page.content)


def get_headers(page: WikipediaPage) -> list:
    headers = []
    body = page.content
    while body.find("== ") > -1:
        header = body[body.find("== ") + 3: body.find(" ==")].strip()
        headers.append(header)
        body = body[body.find("==\n") + 3:].strip()

    return headers


def find_body_of_header(page: WikipediaPage, header: str) -> str:
    body = page.content[page.content.find(header) + len(header):]
    return body[body.find("==\n") + 3:body.find("\n==")].strip()


def format_title(title: str):
    new_title = ""
    while title.find(" ") > -1 or title.find("-") > -1:
        if (title.find(" ") > -1 and title.find("-") == -1) or (title.find(" ") < title.find("-")):
            new_title += title[0].upper() + title[1:title.find(" ")] + " "
            title = title[title.find(" ") + 1:]
        else:
            new_title += title[0].upper() + title[1:title.find("-")] + " "
            title = title[title.find("-") + 1:]

        if len(new_title) > 30:
            break

    if len(new_title) < 30:
        new_title += title[0].upper() + title[1:]

    return new_title


def main():
    name = "Rochester Institute of Technology"
    # codecs.open allows utf-8 for encoding of non-ascii characters
    article_file = codecs.open(name + ".txt", "w", "utf-8")
    page = wiki_search(name)
    write_article_to_text(page, article_file)
    headers = get_headers(page)

    # creates an object of the presentation class, a blank presentation in the user's drive, and a reference to that
    # presentation called "presentation"
    pres_obj = presentation(get_credentials())
    pres = pres_obj.create_presentation(page.title)

    print(pres_obj.set_slide_title(pres.get('presentationId'), pres.get('slides')[0]
                                   .get('pageElements')[0].get('objectId'), format_title(page.title)))
    print(pres_obj.set_slide_title(pres.get('presentationId'), pres.get('slides')[0]
                                   .get('pageElements')[1].get('objectId'), "From Wikipedia"))

    for header in reversed(headers):
        bullet_list = get_bullet_points(page, header, 4)
        if bullet_list.count('\n') < 2:
            print("Skipped slide with only one bullet point:", header)
            continue

        if len(bullet_list) > 0:
            pres_obj.create_slide(pres.get('presentationId'), str(uuid4()), format_title(header), bullet_list)
        else:
            print("Skipped blank header:", header)

    if len(page.summary) > 0:
        pres_obj.create_slide(pres.get('presentationId'), str(uuid4()), "Summary",
                              get_bullet_points(page, "summary", 4))


if __name__ == '__main__':
    main()
