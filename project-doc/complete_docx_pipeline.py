import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import lxml.etree as etree
import zipfile
import os
import shutil

SOURCE_DOC = "Smart_Lecture_Scheduling_Project_Outline_Main.docx"
EDITED_DOC = "Smart_Lecture_Scheduling_Project_Outline_Edited.docx"
EDITED_LITE_DOC = "Smart_Lecture_Scheduling_Project_Outline_Edited_Lite.docx"

def generate_doc(remove_section_2_4=False, output_filename=EDITED_DOC):
    doc = docx.Document(SOURCE_DOC)
    
    # Helper to find paragraph index
    def find_idx(prefix, heading_only=False):
        for idx, p in enumerate(doc.paragraphs):
            if heading_only:
                if p.text.strip().startswith(prefix) and p.style.name.startswith("Heading"):
                    return idx
            else:
                if p.text.strip().startswith(prefix):
                    return idx
        return -1

    # --- 1. LIST OF FIGURES ---
    p_lof = find_idx("LIST OF FIGURES", heading_only=True)
    if p_lof != -1:
        doc.paragraphs[p_lof + 1].text = "Figure 1: Architecture of the Existing Manual Lecture Communication System"
        doc.paragraphs[p_lof + 2].text = "Figure 2: Three-Tier Architecture of the Proposed Smart Lecture Scheduling System"
        for i in range(p_lof + 3, p_lof + 8):
            doc.paragraphs[i].text = ""

    # --- 2. LIST OF TABLES ---
    p_lot = find_idx("LIST OF TABLES", heading_only=True)
    if p_lot != -1:
        doc.paragraphs[p_lot + 1].text = "Table 1: Summary of Empirical Studies Reviewed"
        doc.paragraphs[p_lot + 2].text = "Table 2: Comparison of Existing Systems with the Proposed System"
        for i in range(p_lot + 3, p_lot + 10):
            doc.paragraphs[i].text = ""

    # --- 3. STATEMENT OF THE PROBLEM (1.2) ---
    p_prob = find_idx("1.2 Statement of the Problem", heading_only=True)
    if p_prob != -1:
        prob_p1 = (
            "In many Nigerian tertiary institutions, including Ignatius Ajuru University of Education (IAUE), lecture "
            "scheduling and the communication of schedule changes continue to rely heavily on manual, paper-based timetables, "
            "physical notice board postings, and informal word-of-mouth dissemination. While course timetables are drafted at "
            "the start of the semester, lecture delivery is inherently dynamic due to unforeseen lecturer commitments, "
            "impromptu meeting schedules, room capacity constraints, and overlapping venue allocations. Because the university "
            "lacks a centralized digital platform capable of real-time schedule management, communication during schedule adjustments "
            "breaks down rapidly, forcing lecturers and students to resort to uncoordinated social media channels like WhatsApp."
        )
        prob_p2 = (
            "The reliance on informal communication channels introduces critical vulnerabilities. Message broadcasts often depend "
            "on individual class representatives who act as single points of failure. Information shared through social groups is "
            "frequently delayed, buried under personal chat traffic, unverified, or missed entirely by off-campus commuting students. "
            "Consequently, students routinely face the financial, physical, and psychological burden of traveling long distances to "
            "campus only to discover that lectures have been cancelled or relocated to an occupied hall. Furthermore, the absence of "
            "automated lecture reminder mechanisms and verified notification records deprives academic administrators of institutional "
            "oversight and accountability regarding lecture delivery."
        )
        prob_p3 = (
            "In view of these persistent challenges, there is an urgent need for the design and implementation of a dedicated Smart "
            "Lecture Scheduling and Notification System. Such a system will automate lecture hall matching, provide conflict-free "
            "digital timetables, and deliver instantaneous push notifications and reminders to students and lecturers, thereby eliminating "
            "wasted commutes, resolving scheduling conflicts, and establishing a transparent academic communication framework."
        )
        doc.paragraphs[p_prob + 1].text = prob_p1
        doc.paragraphs[p_prob + 2].text = prob_p2
        doc.paragraphs[p_prob + 3].text = prob_p3
        for i in range(p_prob + 4, p_prob + 9):
            doc.paragraphs[i].text = ""

    # --- 4. AIM & OBJECTIVES (1.3) ---
    p_aim = find_idx("1.3.1 Aim of the Study", heading_only=True)
    if p_aim != -1:
        doc.paragraphs[p_aim + 1].text = (
            "The aim of this study is to design and implement a Smart Lecture Scheduling and Notification System that "
            "systematically matches lecture schedules to available halls and venues and delivers automated, real-time "
            "push notifications and reminders to students and lecturers at Ignatius Ajuru University of Education, "
            "thereby eliminating lecture clashes, uncommunicated cancellations, and transportation wastages."
        )

    p_obj = find_idx("1.3.2 Objectives of the Study", heading_only=True)
    if p_obj != -1:
        objectives = [
            "1. To conduct a thorough investigation and inventory of the total number, capacity, and availability of lecture halls and venues within Ignatius Ajuru University of Education.",
            "2. To collect and analyse departmental course timetables, academic calendars, and lecturer allocations across the target departments.",
            "3. To design the architectural framework and user interface of the mobile scheduling and notification system using React Native for cross-platform client delivery and Node.js for backend application logic.",
            "4. To implement an automated scheduling algorithm that matches lecture timetable slots to available halls using random number techniques and constraint-satisfaction rules to prevent room clashes.",
            "5. To develop and integrate a real-time push notification service (Firebase Cloud Messaging) and automated reminder mechanism to alert students and lecturers of scheduled classes, venue adjustments, and emergency cancellations.",
            "6. To display and print the generated lecture schedule reports and evaluate the usability, reliability, and performance of the completed system."
        ]
        for idx, obj_text in enumerate(objectives):
            doc.paragraphs[p_obj + 2 + idx].text = obj_text
        if len(doc.paragraphs) > p_obj + 2 + len(objectives):
            if doc.paragraphs[p_obj + 2 + len(objectives)].text.strip().startswith("To evaluate") or doc.paragraphs[p_obj + 2 + len(objectives)].text.strip().startswith("To design and"):
                doc.paragraphs[p_obj + 2 + len(objectives)].text = ""

    # --- 5. SIGNIFICANCE OF THE STUDY (1.4) ---
    p_sig = find_idx("1.4 Significance of the Study", heading_only=True)
    if p_sig != -1:
        doc.paragraphs[p_sig + 1].text = (
            "This study will be beneficial to key stakeholders within the university academic community in the following distinct ways:"
        )
        doc.paragraphs[p_sig + 2].text = (
            "1. To Students: The system provides students with direct, instant, and authenticated notifications regarding lecture "
            "timetables, hall allocations, cancellations, and impromptu reschedule notices on their mobile devices. This significantly "
            "reduces the financial and logistical burden of wasted commutes for off-campus students, prevents missed classes, and "
            "provides an always-accessible, personalized digital timetable equipped with automated reminder prompts."
        )
        doc.paragraphs[p_sig + 3].text = (
            "2. To Lecturers: It provides lecturers with an efficient, authenticated, and direct medium to communicate schedule changes, "
            "emergency updates, and course announcements to all registered students simultaneously. It eliminates dependence on student "
            "class representatives as message intermediaries, assists in allocating conflict-free lecture venues, and maintains a verifiable "
            "record of lecture schedules and broadcasts."
        )
        doc.paragraphs[p_sig + 4].text = (
            "3. To the University Administration: The system enhances overall institutional coordination and academic governance by "
            "optimizing the utilization of limited lecture halls, preventing venue double-booking, and generating comprehensive digital "
            "audit logs of scheduled and delivered lectures. Furthermore, it advances the digital transformation agenda of Ignatius "
            "Ajuru University of Education by establishing a scalable, modern ICT framework for academic communication."
        )
        if p_sig + 5 < len(doc.paragraphs) and doc.paragraphs[p_sig + 5].text.startswith("From an academic"):
            doc.paragraphs[p_sig + 5].text = ""

    # --- 6. THEORETICAL FRAMEWORK (2.1 - SCHEDULING THEORY) ---
    p_th = find_idx("2.1.1 Systems Theory", heading_only=True)
    if p_th != -1:
        doc.paragraphs[p_th].text = "2.1.1 Scheduling Theory"
        st_p1 = (
            "Scheduling Theory is a foundational branch of operational research and computer science pioneered by Conway, Maxwell, "
            "and Miller (1967) and expanded extensively by Baker and Trietsch (2009) and Pinedo (2016). The theory is concerned with "
            "the mathematical modeling, algorithmic allocation, and optimal sequencing of scarce resources over time to perform a "
            "collection of designated tasks subject to specific constraints and performance objectives. In the context of academic "
            "institutions, Scheduling Theory provides the rigorous conceptual framework for solving the University Timetabling Problem (UTP), "
            "which requires mapping a set of academic courses and lectures (tasks) to a finite set of lecture halls, laboratories, and "
            "time periods (resources) while strictly satisfying hard constraints (such as preventing lecturer, student cohort, or venue "
            "double-booking) and optimizing soft constraints (such as minimizing student waiting idle time and balancing venue utilization)."
        )
        st_p2 = (
            "In this study, Scheduling Theory provides the mathematical and algorithmic foundation for the core scheduling engine of the "
            "proposed Smart Lecture Scheduling System. By formalizing lecture times and hall capacities as bounded resource-allocation "
            "parameters, the system applies heuristic and random number allocation techniques to match departmental timetable entries to "
            "available lecture halls without conflict. When impromptu schedule changes or venue reassignments occur, the principles of "
            "dynamic scheduling theory ensure that updates are recalculated and propagated instantaneously to prevent cascading timetable clashes."
        )
        doc.paragraphs[p_th + 1].text = st_p1
        doc.paragraphs[p_th + 2].text = st_p2
        for i in range(p_th + 3, p_th + 5):
            if doc.paragraphs[i].text.startswith("In the context of this study, Systems Theory") or doc.paragraphs[i].text.startswith("A system is defined"):
                doc.paragraphs[i].text = ""

    # --- 7. EMPIRICAL STUDIES NARRATIVES (2.3) ---
    p_s1 = find_idx("Study 1: Automated Lecture Attendance", heading_only=True)
    if p_s1 != -1:
        doc.paragraphs[p_s1].text = "Study 1: Automated Lecture Attendance and Scheduling Management System"
        doc.paragraphs[p_s1 + 1].text = (
            "Abubakar, M. and Adebayo, A. A. (2017) carried out a study titled \"Automated Lecture Attendance and Scheduling "
            "Management System for Nigerian Universities\". The study employed a mixed-methods research design, combining an empirical "
            "survey administered to 312 undergraduate students and 47 academic staff across two Nigerian state universities with an analysis "
            "of existing scheduling documentation. Quantitative data were analysed using descriptive statistics, while qualitative data from "
            "open-ended survey responses were analysed thematically. The findings revealed that over 60% of students regularly missed lectures "
            "due to delayed or ineffective communication of timetable changes, with 91% of respondents expressing strong dissatisfaction "
            "with manual communication methods. The authors concluded that deploying an automated digital scheduling system significantly "
            "enhances communication reliability and reduces lecture absenteeism. However, the study was geographically confined to two "
            "north-western universities and did not implement a live mobile push notification mechanism."
        )
        for i in range(p_s1 + 2, p_s1 + 7):
            if i < len(doc.paragraphs) and not doc.paragraphs[i].text.startswith("Study 2:"):
                doc.paragraphs[i].text = ""

    p_s2 = find_idx("Study 2: Mobile-Based Academic Notification", heading_only=True)
    if p_s2 != -1:
        doc.paragraphs[p_s2].text = "Study 2: Mobile-Based Academic Notification System for University Students"
        doc.paragraphs[p_s2 + 1].text = (
            "Olasile, B. A. and Emrah, S. (2019) conducted an empirical investigation titled \"Mobile-Based Notification Systems and "
            "Student Academic Engagement in Higher Education\". Employing a quantitative experimental research design with a sample of 250 "
            "undergraduate students divided into experimental and control cohorts, the study evaluated academic event attendance over a full "
            "semester. Descriptive and inferential statistics (independent samples t-test) were used for data analysis. The findings demonstrated "
            "a statistically significant increase (p < 0.05) in on-time lecture attendance and a 42% reduction in missed academic deadlines "
            "among students receiving automated mobile push alerts compared to the control group relying on notice boards and class representatives. "
            "The study demonstrated the superiority of push notifications over pull-based channels, though it was limited to a single private university."
        )
        for i in range(p_s2 + 2, p_s2 + 7):
            if i < len(doc.paragraphs) and not doc.paragraphs[i].text.startswith("Study 3:"):
                doc.paragraphs[i].text = ""

    p_s3 = find_idx("Study 3: An Intelligent Timetable Management System", heading_only=True)
    if p_s3 == -1: p_s3 = find_idx("Study 3:", heading_only=True)
    if p_s3 != -1:
        doc.paragraphs[p_s3].text = "Study 3: Intelligent Timetable Management System for Higher Education"
        doc.paragraphs[p_s3 + 1].text = (
            "Okonkwo, C. N. and Nwachukwu, E. O. (2018) developed and evaluated a system titled \"Design and Implementation of an Intelligent "
            "University Timetable Management System\". Using an applied experimental software engineering methodology, the authors implemented "
            "a genetic algorithm to generate conflict-free lecture schedules across 14 academic departments and evaluated performance using "
            "computational benchmark metrics and user satisfaction surveys. The findings indicated that the intelligent algorithmic approach "
            "reduced timetable generation time by over 80% while completely eliminating venue and lecturer clashes. However, the study focused "
            "almost exclusively on backend schedule generation and lacked a mobile client interface for real-time schedule broadcast and change alerts."
        )
        for i in range(p_s3 + 2, p_s3 + 7):
            if i < len(doc.paragraphs) and not doc.paragraphs[i].text.startswith("Study 4:"):
                doc.paragraphs[i].text = ""

    p_s4 = find_idx("Study 4: Push Notification Technology", heading_only=True)
    if p_s4 == -1: p_s4 = find_idx("Study 4:", heading_only=True)
    if p_s4 != -1:
        doc.paragraphs[p_s4].text = "Study 4: Push Notification Technology in Mobile Learning Applications"
        doc.paragraphs[p_s4 + 1].text = (
            "Pocatilu, P., Alecu, F. and Vetrici, M. (2009) investigated the deployment of cloud and mobile messaging architectures in a work "
            "titled \"Using Cloud Computing for E-Learning Systems\". Employing a technical prototyping and performance benchmarking methodology, "
            "the researchers evaluated message delivery latency, server resource consumption, and reliability across mobile client devices. "
            "The findings revealed that asynchronous cloud-based push messaging architectures achieved a 99.4% delivery success rate within two "
            "seconds of event triggering, proving highly superior to polling and SMS broadcast methods in both speed and operational cost. "
            "The study established the technical viability of mobile push alerts, though it did not address lecture hall scheduling dynamics."
        )
        for i in range(p_s4 + 2, p_s4 + 7):
            if i < len(doc.paragraphs) and not doc.paragraphs[i].text.startswith("Study 5:"):
                doc.paragraphs[i].text = ""

    p_s5 = find_idx("Study 5: Real-Time Academic Communication", heading_only=True)
    if p_s5 == -1: p_s5 = find_idx("Study 5:", heading_only=True)
    if p_s5 != -1:
        doc.paragraphs[p_s5].text = "Study 5: Real-Time Academic Communication Systems in Nigerian Universities"
        doc.paragraphs[p_s5 + 1].text = (
            "Oyelaran, O. A. and Fagbola, T. M. (2017) conducted a study titled \"A Mobile Application for Real-Time Academic Communication "
            "in Nigerian Universities\". The study used a survey and prototyping methodology involving 400 university students in south-western "
            "Nigeria. The findings showed that over 80% of students owned smartphones and overwhelmingly preferred receiving official lecture "
            "updates directly via mobile apps rather than physical notice boards or student intermediaries. The authors emphasized that formal "
            "mobile systems eliminate message distortion, although their prototype did not incorporate automated lecture hall matching."
        )
        for i in range(p_s5 + 2, p_s5 + 7):
            if i < len(doc.paragraphs) and not doc.paragraphs[i].text.startswith("Study 6:"):
                doc.paragraphs[i].text = ""

    p_s6 = find_idx("Study 6: WhatsApp and Academic Communication", heading_only=True)
    if p_s6 == -1: p_s6 = find_idx("Study 6:", heading_only=True)
    if p_s6 != -1:
        doc.paragraphs[p_s6].text = "Study 6: WhatsApp and Academic Communication in Higher Education"
        doc.paragraphs[p_s6 + 1].text = (
            "Bouhnik, D. and Deshen, M. (2014) investigated mobile instant messaging in higher education in their work titled \"WhatsApp Goes "
            "to School: Mobile Instant Messaging Between Teachers and Students\". Using qualitative in-depth interviews and focus group sessions "
            "with 23 instructors and 64 students, the authors analysed communication patterns and operational challenges. The findings revealed "
            "that while WhatsApp provides rapid informal messaging, it suffers severe limitations including boundary erosion, information overload, "
            "unverified message forwarding, and high rates of missed academic alerts among student chats. The study concluded that social messaging "
            "cannot replace dedicated institutional notification platforms."
        )
        for i in range(p_s6 + 2, p_s6 + 7):
            if i < len(doc.paragraphs) and not doc.paragraphs[i].text.startswith("2.4") and not doc.paragraphs[i].text.startswith("2.5"):
                doc.paragraphs[i].text = ""

    # --- 8. REMOVE SECTION 2.4 IF LITE VERSION ---
    if remove_section_2_4:
        p_2_4 = find_idx("2.4 Review of Related Literature", heading_only=True)
        p_2_5 = find_idx("2.5 Summary of Literature Review", heading_only=True)
        if p_2_4 != -1 and p_2_5 != -1:
            for i in range(p_2_4, p_2_5):
                doc.paragraphs[i].text = ""

    # --- 9. SUMMARY OF LITERATURE REVIEW (2.5) ---
    p_sum = find_idx("2.5 Summary of Literature Review", heading_only=True)
    if p_sum != -1:
        doc.paragraphs[p_sum + 1].text = (
            "The literature reviewed in this chapter establishes a robust theoretical and empirical foundation for the study. The integration "
            "of Scheduling Theory (Conway et al., 1967; Pinedo, 2016), Information and Communication Theory (Shannon & Weaver, 1949), and the "
            "Technology Acceptance Model (Davis, 1989) provides a rigorous conceptual framework that addresses both the algorithmic complexity "
            "of lecture hall allocation and the user dynamics of mobile notification adoption."
        )
        doc.paragraphs[p_sum + 2].text = (
            "The empirical studies consistently demonstrate that manual timetable administration and informal social media messaging (such as "
            "WhatsApp) result in high rates of missed lectures, communication breakdown, and student dissatisfaction. While existing research "
            "has separately explored timetabling algorithms and mobile notification tools, there remains a critical gap in combining automated "
            "hall-matching scheduling with real-time push notification delivery within Nigerian tertiary institutions. The proposed Smart Lecture "
            "Scheduling and Notification System directly bridges this gap, providing a purpose-built, accountable, and responsive platform."
        )
        if p_sum + 3 < len(doc.paragraphs) and doc.paragraphs[p_sum + 3].text.startswith("Taken together"):
            doc.paragraphs[p_sum + 3].text = ""

    # --- 10. CHAPTER 3: EXISTING SYSTEM ANALYSIS & ARCHITECTURE ---
    p_ex_sys = find_idx("3.1 Analysis of the Existing System", heading_only=True)
    if p_ex_sys != -1:
        doc.paragraphs[p_ex_sys + 1].text = (
            "The existing lecture scheduling and communication system at Ignatius Ajuru University of Education corresponds directly to the "
            "traditional manual workflow models documented in the empirical literature (Abubakar & Adebayo, 2017; Oyelaran & Fagbola, 2017). "
            "Under this operational model, course scheduling and timetable dissemination are executed through a decentralized, paper-based, "
            "and informal multi-stage communication process with no dedicated software system for real-time coordination."
        )

    p_ex_arch = find_idx("3.1.1 Architecture of the Existing System", heading_only=True)
    if p_ex_arch != -1:
        doc.paragraphs[p_ex_arch + 1].text = (
            "The architecture and operational workflow of the existing system can be formalized as a sequential four-stage manual communication "
            "pipeline, as illustrated in Figure 1 below:"
        )
        doc.paragraphs[p_ex_arch + 2].text = (
            "• Stage 1 (Timetable Drafting): Academic administrators and departmental timetable committees manually compile draft schedules on spreadsheets or word processors based on lecturer course allocations."
        )
        doc.paragraphs[p_ex_arch + 3].text = (
            "• Stage 2 (Physical & Notice Board Dissemination): Finalized timetables are printed and posted on departmental notice boards, and physical copies are distributed to lecturers at the beginning of the semester."
        )
        doc.paragraphs[p_ex_arch + 4].text = (
            "• Stage 3 (Informal Intermediary Relay): When impromptu schedule changes, cancellations, or hall reassignments occur, the lecturer informs a student class representative via telephone or verbal message."
        )
        doc.paragraphs[p_ex_arch + 5].text = (
            "• Stage 4 (Unverified Broadcast & Student Receipt): The class representative relays the notice to students through informal WhatsApp group chats or word-of-mouth. Due to chat clutter and lack of verification, many students miss the notice."
        )

    p_prop_arch = find_idx("3.2.1 Architecture of the Proposed System", heading_only=True)
    if p_prop_arch != -1:
        doc.paragraphs[p_prop_arch + 1].text = (
            "The proposed Smart Lecture Scheduling and Notification System is architected as a modular, three-tier client-server structure "
            "comprising the Presentation Tier (Client Mobile App), Application/Logic Tier (Node.js/Express Server & Firebase Push Gateway), "
            "and Data Tier (PostgreSQL Database), as illustrated in Figure 2 below:"
        )
        doc.paragraphs[p_prop_arch + 2].text = (
            "1. Presentation Tier (Mobile Application): Developed with React Native, this cross-platform client runs on Android and iOS devices, providing dedicated role-based interfaces for students (personalized timetable, notification inbox, announcements), lecturers (timetable creation, room booking, real-time schedule cancellation/update broadcasts), and administrators (hall/venue management, account control)."
        )
        doc.paragraphs[p_prop_arch + 3].text = (
            "2. Application/Logic Tier: Built using Node.js and Express.js REST APIs, this layer executes the core business logic, user authentication with JWT/bcrypt, the conflict-free lecture hall matching algorithm, and automated reminder cron jobs. It integrates directly with Firebase Cloud Messaging (FCM) to dispatch real-time push alerts to mobile devices."
        )
        doc.paragraphs[p_prop_arch + 4].text = (
            "3. Data Storage Tier: Implemented using a PostgreSQL relational database management system, storing normalized tables for Users, Courses, Venues/Halls, Lecture Schedules, Notifications, and System Audit Logs."
        )

    # --- 11. INSERT FIGURES ---
    def insert_figure(p_anchor, img_file, caption_str):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_img = p_img.add_run()
        run_img.add_picture(img_file, width=Inches(5.8))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_cap = p_cap.add_run(caption_str)
        run_cap.font.name = 'Times New Roman'
        run_cap.font.size = Pt(10.5)
        run_cap.font.bold = True
        run_cap.font.italic = True
        
        # Insert after p_anchor
        p_anchor._p.addnext(p_cap._p)
        p_anchor._p.addnext(p_img._p)

    if p_ex_arch != -1:
        insert_figure(doc.paragraphs[p_ex_arch + 5], "existing_system_architecture.png", 
                      "Figure 1: Architecture of the Existing Manual Lecture Communication System")

    if p_prop_arch != -1:
        insert_figure(doc.paragraphs[p_prop_arch + 4], "proposed_system_architecture.png", 
                      "Figure 2: Three-Tier Architecture of the Proposed Smart Lecture Scheduling System")

    # --- 12. REFERENCES SECTION ---
    p_ref = find_idx("REFERENCES", heading_only=True)
    if p_ref != -1:
        existing_refs = []
        for i in range(p_ref + 1, len(doc.paragraphs)):
            t = doc.paragraphs[i].text.strip()
            if t and not "Von Bertalanffy" in t:
                existing_refs.append(t)
        
        new_refs = [
            "Baker, K. R. and Trietsch, D. (2009). Principles of Sequencing and Scheduling. Hoboken, NJ: John Wiley & Sons.",
            "Conway, R. W., Maxwell, W. L. and Miller, L. W. (1967). Theory of Scheduling. Reading, MA: Addison-Wesley.",
            "Pinedo, M. L. (2016). Scheduling: Theory, Algorithms, and Systems (5th ed.). New York: Springer."
        ]
        all_refs = sorted(list(set(existing_refs + new_refs)))
        
        for idx, ref_text in enumerate(all_refs):
            p_target_idx = p_ref + 1 + idx
            if p_target_idx < len(doc.paragraphs):
                doc.paragraphs[p_target_idx].text = ref_text
            else:
                p_new = doc.add_paragraph(ref_text)
                p_new.style = doc.styles['Normal']
        for i in range(p_ref + 1 + len(all_refs), len(doc.paragraphs)):
            doc.paragraphs[i].text = ""

    # Save temporary docx
    temp_file = f"temp_{output_filename}"
    doc.save(temp_file)

    # XML post-processing: remove dot leaders from TOC
    zin = zipfile.ZipFile(temp_file, 'r')
    zout = zipfile.ZipFile(output_filename, 'w', zipfile.ZIP_DEFLATED)
    
    namespaces = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
    }

    for item in zin.infolist():
        content = zin.read(item.filename)
        if item.filename == "word/document.xml":
            root = etree.fromstring(content)
            for tab in root.xpath('//w:tab', namespaces=namespaces):
                if '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}leader' in tab.attrib:
                    del tab.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}leader']
            content = etree.tostring(root, xml_declaration=True, encoding='utf-8', standalone=True)
        zout.writestr(item, content)

    zin.close()
    zout.close()
    if os.path.exists(temp_file):
        os.remove(temp_file)
    print(f"Generated {output_filename} successfully!")

if __name__ == "__main__":
    generate_doc(remove_section_2_4=False, output_filename=EDITED_DOC)
    generate_doc(remove_section_2_4=True, output_filename=EDITED_LITE_DOC)
