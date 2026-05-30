import io
import os
import tempfile
import unittest

from openpyxl import Workbook

from app import create_app, init_db


class LifecycleAppTest(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.database = os.path.join(self.tempdir.name, "test.db")
        init_db(self.database)
        self.app = create_app({"TESTING": True, "DATABASE": self.database, "SECRET_KEY": "test"})
        self.client = self.app.test_client()

    def tearDown(self):
        self.tempdir.cleanup()

    def test_seeded_options_include_required_statuses_and_stages(self):
        response = self.client.get("/api/options")
        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        statuses = {item["value"] for item in payload["status"]}
        stages = {item["value"] for item in payload["stage"]}
        self.assertTrue({"Not Started", "In Progress", "On Hold", "Completed"}.issubset(statuses))
        self.assertTrue({"I20", "I50", "D20", "Qualified Out", "I0", "D0", "D50"}.issubset(stages))

    def test_record_crud_and_validation(self):
        bad_response = self.client.post(
            "/api/records",
            json={"trackerType": "I20", "title": "Bad date", "stage": "I20", "status": "Not Started", "owner": "Owner", "dueDate": "bad"},
        )
        self.assertEqual(bad_response.status_code, 400)
        self.assertIn("valid date", " ".join(bad_response.get_json()["errors"]))

        create_response = self.client.post(
            "/api/records",
            json={
                "trackerType": "I20",
                "title": "Lifecycle opportunity",
                "stage": "I20",
                "status": "In Progress",
                "owner": "Owner",
                "customer": "Customer",
                "dueDate": "2030-01-15",
                "priority": "High",
            },
        )
        self.assertEqual(create_response.status_code, 201)
        record_id = create_response.get_json()["id"]

        list_response = self.client.get("/api/records?trackerType=I20&status=In%20Progress")
        records = list_response.get_json()["records"]
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["title"], "Lifecycle opportunity")

        update_response = self.client.put(
            f"/api/records/{record_id}",
            json={
                "trackerType": "I20",
                "title": "Lifecycle opportunity",
                "stage": "I50",
                "status": "Completed",
                "owner": "Owner",
                "customer": "Customer",
                "dueDate": "2030-01-15",
                "priority": "High",
            },
        )
        self.assertEqual(update_response.status_code, 200)

        detail_response = self.client.get(f"/api/records/{record_id}")
        detail = detail_response.get_json()
        self.assertEqual(detail["record"]["status"], "Completed")
        self.assertTrue(detail["record"]["completedDate"])
        self.assertGreaterEqual(len(detail["history"]), 1)

    def test_csv_import_preserves_unknown_columns(self):
        csv_bytes = b"Opportunity Name,Owner,Stage,Status,Due Date,Workbook Extra\nImported item,Owner,I20,Not Started,2030-02-01,Preserved\n"
        response = self.client.post(
            "/api/import",
            data={"trackerType": "I20", "file": (io.BytesIO(csv_bytes), "records.csv")},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["imported"], 1)

        records = self.client.get("/api/records?trackerType=I20").get_json()["records"]
        self.assertEqual(records[0]["additionalFields"]["Workbook Extra"], "Preserved")

    def test_lifecycle_workbook_import_uses_tracker_sheet_and_fields(self):
        workbook = Workbook()
        overview = workbook.active
        overview.title = "Overview"
        overview.append(["This sheet should not import"])
        tracker = workbook.create_sheet("D20 Tracker")
        tracker.append([
            "Stage",
            "Activity",
            "Status",
            "who @ cortave ",
            "Start Day",
            "Actual Start Date \n(Edit Kick Off Date for Automated deadlines)",
            "Target Due Day",
            "Due Date",
            "cortave Owner",
            "Innovator Owner",
            "Link",
            "Notes",
        ])
        tracker.append([
            "D20",
            "Discovery meeting",
            "In Progress",
            "Wade / Sales",
            20,
            "2030-01-01",
            10,
            "2030-01-11",
            "Bobby",
            "Innovator contact",
            "https://example.com",
            "Workbook note",
        ])

        org = workbook.create_sheet("Org Impact ")
        org["C4"] = 5
        org["D4"] = "Working days per week"
        org["C10"] = 2008
        org["D10"] = "Working hours per year (average)"
        org["C11"] = 100000
        org["D11"] = "Average FTE cost (annual)"
        org["C16"] = "Partner Manager"
        org["D16"] = 15
        org["E16"] = 45
        org["F16"] = 90
        org["G16"] = 3
        org["H16"] = 9
        org["I16"] = 18
        org["C26"] = "Total"
        org["D26"] = 15
        org["E26"] = 45
        org["F26"] = 90
        org["G26"] = 3
        org["H26"] = 9
        org["I26"] = 18
        org["C31"] = "Partner Manager"
        org["D31"] = 120
        org["E31"] = 360
        org["F31"] = 720
        org["G31"] = 5976.1
        org["H31"] = 17928.3
        org["I31"] = 35856.6
        org["C41"] = "Total"
        org["D41"] = 120
        org["E41"] = 360
        org["F41"] = 720
        org["G41"] = 5976.1
        org["H41"] = 17928.3
        org["I41"] = 35856.6
        org["C45"] = "Recruitment"
        org["D45"] = 40
        org["E45"] = 7360
        org["F45"] = 920
        org["G45"] = 3.67
        org["C49"] = "Total"
        org["D49"] = 49
        org["E49"] = 14856
        org["F49"] = 1857
        org["G49"] = 7.4

        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)

        response = self.client.post(
            "/api/import",
            data={"trackerType": "D20", "file": (output, "Lifecycle Workbook.xlsx")},
            content_type="multipart/form-data",
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["imported"], 1)
        records = self.client.get("/api/records?trackerType=D20").get_json()["records"]
        self.assertEqual(records[0]["title"], "Discovery meeting")
        self.assertEqual(records[0]["whoAtCortave"], "Wade / Sales")
        self.assertEqual(records[0]["startDay"], "20")
        self.assertEqual(records[0]["actualStartDate"], "2030-01-01")
        self.assertEqual(records[0]["targetDueDay"], "10")
        self.assertEqual(records[0]["dueDate"], "2030-01-11")
        self.assertEqual(records[0]["cortaveOwner"], "Bobby")
        self.assertEqual(records[0]["innovatorOwner"], "Innovator contact")
        self.assertEqual(records[0]["link"], "https://example.com")

    def test_partner_dashboard_import_upsert_and_linking(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Partners"
        sheet.append([
            "Partner",
            " Owner ",
            "Territory",
            "Active",
            "Date Of Stage Change",
            "Age of Stage",
            "Days Overdue",
            "Workbook Link",
            "SF Account",
            "Next Steps/notes",
            "Partner Type ",
            "CSM Involved ",
        ])
        sheet.append([
            "Deloitte",
            "Mark",
            "EMEA",
            "Onboarding",
            "2023-06-01",
            "#NAME?",
            12,
            "Deloitte Workbook",
            "Existing account",
            "Talk to account team",
            "Martech",
            "No",
        ])
        sheet.append([
            "Deloitte",
            "Melissa",
            "Global",
            "Recruitment",
            "2023-07-01",
            4,
            0,
            "https://example.com/deloitte",
            "Updated account",
            "Updated notes",
            "Technology",
            "Yes",
        ])
        org = workbook.create_sheet("Org Impact ")
        org["C4"] = 5
        org["D4"] = "Working days per week"
        org["C10"] = 2008
        org["D10"] = "Working hours per year (average)"
        org["C11"] = 100000
        org["D11"] = "Average FTE cost (annual)"
        org["C16"] = "Partner Manager"
        org["D16"] = 15
        org["E16"] = 45
        org["F16"] = 90
        org["G16"] = 3
        org["H16"] = 9
        org["I16"] = 18
        org["C26"] = "Total"
        org["D26"] = 15
        org["E26"] = 45
        org["F26"] = 90
        org["G26"] = 3
        org["H26"] = 9
        org["I26"] = 18
        org["C31"] = "Partner Manager"
        org["D31"] = 120
        org["E31"] = 360
        org["F31"] = 720
        org["G31"] = 5976.1
        org["H31"] = 17928.3
        org["I31"] = 35856.6
        org["C41"] = "Total"
        org["D41"] = 120
        org["E41"] = 360
        org["F41"] = 720
        org["G41"] = 5976.1
        org["H41"] = 17928.3
        org["I41"] = 35856.6
        org["C45"] = "Recruitment"
        org["D45"] = 40
        org["E45"] = 7360
        org["F45"] = 920
        org["G45"] = 3.67
        org["C49"] = "Total"
        org["D49"] = 49
        org["E49"] = 14856
        org["F49"] = 1857
        org["G49"] = 7.4

        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)

        response = self.client.post(
            "/api/partners/import",
            data={"file": (output, "Partner Dashboard.xlsx")},
            content_type="multipart/form-data",
        )

        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertEqual(payload["created"], 1)
        self.assertEqual(payload["updated"], 1)
        self.assertEqual(payload["errors"], [])

        partners_response = self.client.get("/api/partners?search=Deloitte")
        partners = partners_response.get_json()["partners"]
        self.assertEqual(len(partners), 1)
        self.assertEqual(partners[0]["owner"], "Melissa")
        self.assertEqual(partners[0]["masterStage"], "Recruitment")
        self.assertEqual(partners[0]["workbookLink"], "https://example.com/deloitte")
        self.assertTrue(partners[0]["isWorkbookLinkUrl"])

        org_response = self.client.get("/api/org-impact")
        org_payload = org_response.get_json()
        self.assertEqual(org_payload["totals"]["totalHours"], 14856)
        self.assertEqual(org_payload["summary"][0]["label"], "Recruitment")
        self.assertEqual(org_payload["assumptions"][0]["label"], "Partner Manager")

        link_response = self.client.post("/api/lifecycle-workbook/link", json={"partnerId": partners[0]["id"]})
        self.assertEqual(link_response.status_code, 200)
        workbook_response = self.client.get("/api/lifecycle-workbook")
        linked = workbook_response.get_json()["workbook"]
        self.assertEqual(linked["partner"]["partnerName"], "Deloitte")
        self.assertEqual(linked["partner"]["owner"], "Melissa")

        unlink_response = self.client.post("/api/lifecycle-workbook/unlink", json={})
        self.assertEqual(unlink_response.status_code, 200)
        self.assertIsNone(unlink_response.get_json()["workbook"]["partner"])

    def test_partner_dashboard_requires_partners_sheet(self):
        workbook = Workbook()
        workbook.active.title = "Wrong Sheet"
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)

        response = self.client.post(
            "/api/partners/import",
            data={"file": (output, "Partner Dashboard.xlsx")},
            content_type="multipart/form-data",
        )

        self.assertEqual(response.status_code, 400)
        self.assertIn("Partners sheet", response.get_json()["error"])


if __name__ == "__main__":
    unittest.main()
