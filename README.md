# CV Builder

One source file, a tailored CV for every job ad.

I kept rewriting my CV by hand for each application. Now every fact lives once in `cv_data.md`, and each job ad gets a short `@version` block that picks the title, summary, jobs and education it needs. One command turns that into a clean `.docx`.

## How it works

1. Write your facts once in `cv_data.md`: contact, jobs, education, skills.
2. For a new job ad, add a version block:
   ```
   ### @version acme
   title: Program Manager
   summary: Two sentences written for this role.
   jobs: acme, studio
   education: msc
   skills: yes
   ```
3. Build it:
   ```bash
   pip install python-docx
   python3 cv_build.py            # list versions
   python3 cv_build.py acme       # build CV_acme.docx
   python3 cv_build.py acme --photo
   ```
4. Open the file from `output/` in Pages or Word and export to PDF.

**Bold** in the data file is stripped in the output, so you can mark key words for yourself while editing.

For the photo option, put a `headshot.png` next to the script.

## Why this way

* Facts never drift between versions, because there is only one copy.
* A new application takes minutes: write a summary, pick the blocks, run.
* Plain text source works well with AI: paste a job ad and ask for a version block.

All data in `cv_data.md` is example data.
