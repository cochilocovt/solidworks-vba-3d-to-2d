"""ISelectData.View binding must never abort a transaction.

Setting ISelectData.View raises runtime error 91 in some drawing contexts on
this build. Modules 13, 15 and 16 each wrapped that assignment; Module4 did
not, so every dimension-arrange attempt in the 18:45, 23:24 and 23:43
production runs died in its error handler. Until r40 the handler also
destroyed the error number, so the failure read "Dimension arrange API error
in 'Drawing View1': 0: " and could not be diagnosed from its own evidence.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"

MODULES_THAT_BIND_SELECT_DATA_VIEW = (
    "Module4_ModelItemImporter.bas",
    "Module13_ProjectionResolution.bas",
    "Module15_OrdinateScheme.bas",
    "Module16_CalloutDefinition.bas",
)


class R23SelectDataBindingContracts(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="cp1252")

    def code(self, name: str) -> str:
        return "\n".join(
            line for line in self.read(name).split("\n")
            if not line.lstrip().startswith("'")
        )

    def test_every_binder_is_wrapped(self):
        for module in MODULES_THAT_BIND_SELECT_DATA_VIEW:
            with self.subTest(module=module):
                body = self.code(module)
                self.assertIn("Private Function TryBindSelectDataView(", body)
                # Exactly one assignment, and it lives inside the wrapper.
                self.assertEqual(body.count("Set selectData.View = swView"), 1)

    def test_no_module_assigns_the_property_outside_the_wrapper(self):
        """The raw assignment is what aborted Module4 on every run."""
        for module in MODULES_THAT_BIND_SELECT_DATA_VIEW:
            with self.subTest(module=module):
                body = self.code(module)
                wrapper = body.split(
                    "Private Function TryBindSelectDataView("
                )[1].split("\nEnd Function")[0]
                self.assertIn("Set selectData.View = swView", wrapper)
                outside = body.replace(wrapper, "")
                self.assertNotIn("Set selectData.View", outside)

    def test_binder_reports_the_error_number_rather_than_raising(self):
        for module in MODULES_THAT_BIND_SELECT_DATA_VIEW:
            with self.subTest(module=module):
                wrapper = self.code(module).split(
                    "Private Function TryBindSelectDataView("
                )[1].split("\nEnd Function")[0]
                self.assertIn('"Bound"', wrapper)
                self.assertIn('"UnboundAfterError:"', wrapper)
                self.assertIn("CStr(Err.Number)", wrapper)

    def test_arrange_records_the_binding_and_continues(self):
        """Binding scopes the selection; it is not a precondition. The view
        is already activated and read back before this point."""
        body = self.code("Module4_ModelItemImporter.bas")
        self.assertIn("DIMENSION_ARRANGE_BINDING|view=", body)
        self.assertIn("viewBinding = TryBindSelectDataView(", body)

        arrange = body.split("viewBinding = TryBindSelectDataView(")[1]
        # An unbound view must not short-circuit the transaction.
        self.assertNotIn(
            'arrangeOutcome = "FAILED_VIEW_BINDING"',
            arrange.split("SafeExit:")[0],
        )
        self.assertIn("Select3(", arrange)


if __name__ == "__main__":
    unittest.main()
