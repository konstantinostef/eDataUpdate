import re
import csv
from io import StringIO

INPUT_FILE = "./files/teachers_insert_geniki.sql"
OUTPUT_FILE = "./files/teachers_insert_geniki_sqlserver.sql"
NEXT_TEACHER_ID = 10000

def split_values(values):
    return next(csv.reader(StringIO(values), delimiter=',', quotechar="'", skipinitialspace=True))


with open(INPUT_FILE, "r", encoding="utf-8") as f:
    sql = f.read()

pattern = re.compile(
    r"INSERT INTO `teachers`\s*\((.*?)\)\s*VALUES\s*\((.*?)\)\s*ON DUPLICATE KEY UPDATE",
    re.S,
)

output = ["USE [ePyspe]\nGO\n"]

for m in pattern.finditer(sql):
    teacher_id = NEXT_TEACHER_ID
    NEXT_TEACHER_ID += 1

    columns = [c.strip(" `") for c in m.group(1).split(",")]
    values = split_values(m.group(2))

    row = dict(zip(columns, values))

    afm = row["afm"]
    name = row["name"]
    surname = row["surname"]
    fname = row["fname"] or ""
    mname = row["mname"] or ""
    phone = row["telephone"]
    email = "NULL" if row["mail"].upper() == "NULL" else f"'{row['mail']}'"
    schmail = "NULL" if row["sch_mail"].upper() == "NULL" else f"'{row['sch_mail']}'"
    branch = row["klados"]
    am = row["am"]
    employment = row["sxesi_ergasias_id"]
    org_eae = row["org_eae"]
    organic_school = row["organiki_id"]
    active = row["active"]
    director = row["is_director"]
    subdirector = row["is_subdirector"]
    created = row["created_at"]
    updated = row["updated_at"]

    output.append(f"""
IF EXISTS (SELECT 1 FROM [lut].[Teacher] WHERE [AFM] = '{afm}')
BEGIN
    UPDATE [lut].[Teacher]
    SET
        [FirstName] = N'{name}',
        [LastName] = N'{surname}',
        [FathersName] = N'{fname}',
        [MothersName] = N'{mname}',
        [PhoneNo] = '{phone}',
        [Email] = {email},
        [Branch] = N'{branch}',
        [AM] = '{am}',
        [EmploymentRelationshipId] = {employment},
        [OrgEAE] = {org_eae},
        [OrganicSchoolId] = {organic_school},
        [IsDirector] = {director},
        [IsSubdirector] = {subdirector},
        [IsActive] = {active},
        [UpdatedAt] = '{updated}'
    WHERE [AFM] = '{afm}';
END
ELSE
BEGIN
    INSERT INTO [lut].[Teacher]
    (
        [TeacherId],
        [FirstName],
        [LastName],
        [FathersName],
        [MothersName],
        [AFM],
        [AM],
        [Gender],
        [PhoneNo],
        [Email],
        [SchMail],
        [Branch],
        [EmploymentRelationshipId],
        [ServiceTypeId],
        [OrgEAE],
        [OrganicDirectorateId],
        [OrganicSchoolId],
        [ServiceSchoolId],
        [AppointmentDate],
        [AppointmentFEK],
        [IsDirector],
        [IsSubdirector],
        [IsActive],
        [CreatedAt],
        [UpdatedAt]
    )
    VALUES
    (
        {teacher_id},
        N'{name}',
        N'{surname}',
        N'{fname}',
        N'{mname}',
        '{afm}',
        '{am}',
        N'Θ',
        '{phone}',
        {email},
        {schmail},
        N'{branch}',
        {employment},
        NULL,
        {org_eae},
        NULL,
        {organic_school},
        NULL,
        NULL,
        NULL,
        {director},
        {subdirector},
        {active},
        '{created}',
        '{updated}'
    );
END
GO

""")

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    f.write("".join(output))

print("Conversion completed.")