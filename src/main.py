from outlook_reader import get_outlook_emails
from report_generator import generate_word_report
from summarizer import summarize_emails
from pathlib import Path

def main():
    emails = get_outlook_emails(days=7)

    print(f"{len(emails)} e-mails trouvés")

    summary = summarize_emails(emails)
    print(summary)

    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    output_file = output_dir / "compte_rendu_hebdomadaire.docx"
    generate_word_report(emails, str(output_file))

    print(f"Rapport créé : {output_file}")

if __name__ == "__main__":
    main()