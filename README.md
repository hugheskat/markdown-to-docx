## Summary

This will convert markdown files with basic markdown syntax to docx. I wrote this script because Pandoc is blocked by my company. The script is adapted from: 
https://www.cyberforum.ru/python-beginners/thread2607589.html.

It will not convert everything and the docx files are likely to require some editing.

It WILL convert the following markdown elements to their docx equivalent:

- Headings
  - If the first line of the markdown file is not a heading (ie preceded by #), it will be converted to the title heading.
  - Any markdown headings will be converted to the equivalent docx heading level .
- Lists:
  - Unordered lists up to 3 sub-levels.
    - Sub-level 2 list must be indented by 2 spaces, sub-level 3 list by 4 spaces etc.
  - Ordered lists up to 3 sub-levels.
    - Sub-level 2 list must be indented by 2 spaces, sub-level 3 list by 4 spaces etc.
    - Any nested ordered lists will be numbered 1., 2., 3., etc. in the output docx file NOT a), b), c), etc or i., ii., iii.
- Formatting:
  - Bold text
    - Must be denoted as follows in the markdown file:
    ```
    **bold text**
    ```

  - Italic text
    - Must be denoted as follows in the markdown file:
    ```
    *Italic text*
    ```

  - Bold text and Italic text.
    - Must be denoted as follows in the markdown file:
    ```
    ***Bold and italic text***
    ```
    
- Inline code
  - Inline code will be converted to 9pt Consolas font and have light grey highlighting.
  - Must be denoted as follows in the markdown file:
    ```
    `Inline code`
    ```
    
- Horizontal rules
  - Must be denoted as follows in the markdown file:
    ```
    ---
    ```
    
- Block quotes.
  - Will be indented in the output docx file
  - Must be denoted as follows in the markdown file:
    ```
    > blockquote
    >> blockquote
    ```
    
- Source code:
  - Anything, apart from a list item, denoted by a 4 space indent will be indented, highlighted light grey and have
   9pt Consolas font.
   
- Hyperlinks
  - Must be denoted as follows in the markdown file:
     ```
    [Link](https://www.google.com)
    ```

The following elements/syntax WILL NOT be converted:

- Formatting within formatting, e.g. italic text within bold text.
- Images.
- Block code denoted by wrapping in ` ``` `:
  - Block code will ONLY be converted if it's indented by 4 spaces in the markdown.
- Nested lists below sub-level 3.


**Version of Python used for testing: 3.6**

## Input required

- The path to the folder containing the .md files to be converted.

## Output

- The converted .docx files will be output in an 'Output docx' folder within the user-specified folder.
