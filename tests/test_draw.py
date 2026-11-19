import unittest

from draw import DrawConfigurationError, plan_draw


class FamilyDrawTests(unittest.TestCase):
    def setUp(self):
        self.participants = {
            "Ana": "f01",
            "Beto": "f01",
            "Carla": "f02",
            "Diego": "f03",
        }

    def test_random_draw_can_finish_for_every_participant(self):
        assignments = []
        for giver in ("Carla", "Ana", "Diego", "Beto"):
            partner = plan_draw(self.participants, assignments, shuffle=True)[giver]
            self.assertNotEqual(self.participants[giver], self.participants[partner])
            assignments.append({"name": giver, "partner": partner})

        self.assertEqual({item["partner"] for item in assignments}, set(self.participants))
        self.assertEqual(plan_draw(self.participants, assignments), {})

    def test_reserves_recipients_from_existing_draws(self):
        assignments = [
            {"name": "Ana", "partner": "Carla"},
            {"name": "Beto", "partner": "SIN PAREJA (no hay más participantes disponibles)"},
        ]
        plan = plan_draw(self.participants, assignments)
        self.assertNotIn("Ana", plan)
        self.assertNotIn("Carla", plan.values())
        self.assertEqual(plan["Beto"], "Diego")

    def test_rejects_a_draw_that_leaves_someone_without_a_recipient(self):
        with self.assertRaises(DrawConfigurationError):
            plan_draw(self.participants, [{"name": "Carla", "partner": "Diego"}])

    def test_rejects_existing_assignments_inside_a_family(self):
        with self.assertRaises(DrawConfigurationError):
            plan_draw(self.participants, [{"name": "Ana", "partner": "Beto"}])

    def test_requires_a_group_for_each_participant(self):
        self.participants["Ana"] = ""
        with self.assertRaises(DrawConfigurationError):
            plan_draw(self.participants, [])
