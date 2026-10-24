#!/usr/bin/env python3
"""
CV de recherche (1 page) de Sahar Saidani, généré avec reportlab (texte sélectionnable).

Usage :
  python build_cv_research.py                              # CV générique
  python build_cv_research.py --lab "l'équipe Avalon (Inria Lyon)" --out CV_Avalon.pdf
  python build_cv_research.py --skills "Cloud & IaC|AWS (EKS), Terraform" --skills "Langages|Python, Bash"

Seules les informations présentes dans le CV d'origine sont utilisées.
"""
import argparse
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable, KeepTogether

FONT_DIR = "/usr/share/fonts/truetype/liberation/"
pdfmetrics.registerFont(TTFont("LS", FONT_DIR + "LiberationSans-Regular.ttf"))
pdfmetrics.registerFont(TTFont("LS-B", FONT_DIR + "LiberationSans-Bold.ttf"))
pdfmetrics.registerFont(TTFont("LS-I", FONT_DIR + "LiberationSans-Italic.ttf"))
pdfmetrics.registerFontFamily("LS", normal="LS", bold="LS-B", italic="LS-I", boldItalic="LS-B")

NAVY = colors.HexColor("#1F3864")
GREY = colors.HexColor("#444444")

DEFAULT_SKILLS = [
    ("Cloud et infrastructure as code", "AWS (EKS, ECS/Fargate, Lambda), Terraform, Ansible"),
    ("CI/CD et DevSecOps", "GitLab CI/CD, Jenkins, GitHub Actions, ArgoCD, SonarQube, Trivy"),
    ("Conteneurs et supervision", "Docker, Kubernetes, Prometheus, Grafana"),
    ("IA et MLOps", "PyTorch, MLflow, RAG / LLM"),
    ("Langages", "Python, Java, Bash"),
]


def build(out: str, lab: str | None, skills: list[tuple[str, str]], size: float):
    lead = size * 1.22
    base = ParagraphStyle("base", fontName="LS", fontSize=size, leading=lead, textColor=colors.black, alignment=TA_LEFT)
    bullet = ParagraphStyle("bullet", parent=base, leftIndent=9, bulletIndent=1, spaceAfter=0.3)
    h = ParagraphStyle("h", parent=base, fontName="LS-B", fontSize=size + 1.4, textColor=NAVY, spaceBefore=4.5, spaceAfter=0.5)
    name = ParagraphStyle("name", parent=base, fontName="LS-B", fontSize=17, leading=20, textColor=NAVY)
    sub = ParagraphStyle("sub", parent=base, fontSize=size + 0.6, textColor=GREY)
    small = ParagraphStyle("small", parent=base, fontSize=size - 0.2, textColor=GREY)
    role = ParagraphStyle("role", parent=base, spaceBefore=2.2, spaceAfter=0.4)

    doc = SimpleDocTemplate(out, pagesize=A4, leftMargin=14 * mm, rightMargin=14 * mm,
                            topMargin=11 * mm, bottomMargin=9 * mm,
                            title="CV – Sahar Saidani", author="Sahar Saidani")
    s = []

    def section(title):
        s.append(Paragraph(title.upper(), h))
        s.append(HRFlowable(width="100%", thickness=0.6, color=NAVY, spaceAfter=1.8))

    def b(text):
        s.append(Paragraph(text, bullet, bulletText="•"))

    s.append(Paragraph("SAHAR SAIDANI", name))
    s.append(Paragraph("Étudiante ingénieure en génie logiciel | DevOps et cloud | Stage de recherche de 6 mois (février 2027)", sub))
    s.append(Paragraph("sahar.saidani@etudiant-fst.utm.tn | +216 51 480 831 | Tunis, Tunisie", small))
    s.append(Paragraph("linkedin.com/in/sahar-saidani-a209a3261 | github.com/sahar-saidani", small))

    section("Profil")
    target = f" au sein de {lab}" if lab else ""
    s.append(Paragraph(
        "Étudiante ingénieure en génie logiciel (FST Tunis, diplôme 2027), spécialisée en DevOps et cloud "
        "(pipelines CI/CD, déploiement, supervision). Recherche un stage de recherche de 6 mois"
        f"{target}, disponible dès février 2027, en vue d'une thèse.", base))

    section("Compétences")
    for k, v in skills:
        b(f"<b>{k} :</b> {v}")

    section("Expérience professionnelle")
    s.append(Paragraph("<b>Ingénieure DevOps (stage)</b> | VERMEG | Juillet – Août 2026 | Tunis", role))
    b("Développé un microservice Python/FastAPI collectant 5 sources de données DevOps sur 2 environnements (Linux, Windows).")
    b("Conçu une architecture distribuée à 3 composants (Spring Boot, RabbitMQ, PostgreSQL) traitant les flux de 5 agents de collecte en temps réel et en différé.")
    b("Conteneurisé avec Docker, industrialisé le déploiement via GitLab CI/CD et surveillé via Grafana (~10 métriques par type de donnée collectée).")
    b("<b>Résultat :</b> réduit le temps d'accès DevOps de 10 min à 30 s (−95 %).")
    s.append(Paragraph("<b>Ingénieure IA (stage)</b> | GRSS Consulting | Août – Septembre 2025 | Tunis", role))
    b("Élaboré un système de classification et un agent RAG sur 50 documents pédagogiques.")
    b("Implémenté l'extraction, le découpage en chunks (500 tokens, chevauchement de 50) et la recherche de passages.")
    b("Réalisé une application 3 tiers (Next.js, Prisma, MongoDB) intégrant classification et agent RAG.")
    b("Collaboré avec le client en Agile/Scrum (4 sprints, démo mensuelle) pour valider les livrables.")
    b("<b>Résultat :</b> ramené le temps de traitement de 14 jours à ~1 min (−99,995 %).")
    s.append(Paragraph("<b>Scrum Master / Cloud Contributor</b> | Optima Junior Entreprise | Décembre 2024 – Mai 2025 | Tunis", role))
    b("Encadré une équipe de 4 membres sur AgroVal+, plateforme e-commerce MERN.")
    b("Planifié des sprints de 2 semaines, réparti les tâches et suivi l'avancement au quotidien pour prévenir les blocages.")
    b("Déployé le cloud sur AWS (ECS/Fargate, S3, VPC, ALB, ECR, ACM, DynamoDB).")
    b("<b>Résultat :</b> raccourci l'intégration des fonctionnalités de 5 à 2 jours (−60 %).")

    section("Projets")
    s.append(Paragraph("<b>Pipeline CI/CD DevSecOps sur AWS EKS</b>", role))
    b("<b>Défi :</b> automatiser et sécuriser le déploiement continu sur Kubernetes.")
    b("<b>Action :</b> automatisé un pipeline de 8 étapes (Jenkins, Terraform, ArgoCD) intégrant 3 contrôles qualité/sécurité (SonarQube, Trivy, OWASP).")
    b("<b>Résultat :</b> optimisé le cache Docker et des dépendances : −30 % de durée du pipeline.")
    s.append(Paragraph("<b>Infrastructure serverless sur AWS</b>", role))
    b("<b>Défi :</b> provisionner une application serverless de façon reproductible.")
    b("<b>Action :</b> architecturé une solution de 4 services AWS (Amplify, API Gateway, Lambda, DynamoDB).")
    b("<b>Résultat :</b> infrastructure entièrement provisionnée avec Terraform (IaC).")
    s.append(Paragraph("<b>Architecture cloud 3 tiers multi-AZ sur AWS</b>", role))
    b("<b>Défi :</b> concevoir une architecture reproductible et supervisée.")
    b("<b>Action :</b> structuré 6 services (EC2, ALB, Auto Scaling, RDS, S3, CloudFront) dans un VPC (6 subnets, 2 NAT Gateways) et dockerisé les applications pour ECS/Fargate via ECR.")
    b("<b>Résultat :</b> supervision par CloudWatch (alarmes CPU &gt; 70 % + SNS) et CloudTrail ; HTTPS via ACM.")

    section("Formation")
    b("<b>Diplôme d'ingénieur en génie logiciel</b> | Faculté des Sciences de Tunis | 2024 – 2027 (prévu) | Cours : cloud, DevOps, microservices, architectures distribuées")
    b("<b>Cycle préparatoire intégré en mathématiques, physique et informatique (MPI)</b> | Faculté des Sciences de Tunis | 2022 – 2024 | Cours : programmation C et C++")
    b("<b>Baccalauréat en mathématiques</b> | mention très bien")

    section("Certifications et langues")
    b("<b>Certifications :</b> AWS Certified Cloud Practitioner (2026) | NVIDIA Deep Learning Fundamentals | AI &amp; GenAI (Udacity) | Cybersecurity Certification (CCEP)")
    b("<b>Langues :</b> arabe (langue maternelle) | français (C1) | anglais (B2) | allemand (A1)")

    doc.build(s)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="CV_recherche_sahar_saidani.pdf")
    ap.add_argument("--lab", default=None, help="ex. \"l'équipe Avalon (Inria Lyon)\"")
    ap.add_argument("--skills", action="append", help="'Catégorie|liste' ; répétable, remplace les compétences par défaut")
    ap.add_argument("--size", type=float, default=9.3)
    a = ap.parse_args()
    sk = DEFAULT_SKILLS
    if a.skills:
        sk = [tuple(x.split("|", 1)) for x in a.skills]
    build(a.out, a.lab, sk, a.size)
    print("CV généré :", a.out)
