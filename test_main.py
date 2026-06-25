"""
Tests for the Study Hours vs Marks Prediction project.

This file checks the most important program functions. It imports functions
from main.py and verifies that they return safe and expected results.
"""

import unittest

from main import (
    create_student_dataset,
    evaluate_prediction_model,
    predict_student_marks,
    prepare_training_and_testing_data,
    train_prediction_model,
    validate_study_hours,
)


class DatasetTests(unittest.TestCase):
    """Check that the sample dataset has the expected structure."""

    def test_dataset_contains_expected_columns_and_rows(self):
        student_dataset = create_student_dataset()

        self.assertEqual(list(student_dataset.columns), ["Hours", "Marks"])
        self.assertEqual(len(student_dataset), 20)
        self.assertFalse(student_dataset.isnull().any().any())


class StudyHoursValidationTests(unittest.TestCase):
    """Check valid and invalid study-hour inputs."""

    def test_whole_number_is_accepted(self):
        self.assertEqual(validate_study_hours("5"), 5.0)

    def test_decimal_number_is_accepted(self):
        self.assertEqual(validate_study_hours("5.5"), 5.5)

    def test_empty_input_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "cannot be empty"):
            validate_study_hours("   ")

    def test_text_input_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "must be a number"):
            validate_study_hours("five")

    def test_non_finite_number_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "normal number"):
            validate_study_hours("NaN")

    def test_negative_input_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "cannot be negative"):
            validate_study_hours("-1")

    def test_more_than_24_hours_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "cannot be more than 24"):
            validate_study_hours("25")


class PredictionModelTests(unittest.TestCase):
    """Check model training, evaluation, and prediction behavior."""

    @classmethod
    def setUpClass(cls):
        cls.student_dataset = create_student_dataset()
        (
            cls.training_hours,
            cls.testing_hours,
            cls.training_marks,
            cls.testing_marks,
        ) = prepare_training_and_testing_data(cls.student_dataset)
        cls.prediction_model = train_prediction_model(
            cls.training_hours,
            cls.training_marks,
        )

    def test_model_evaluation_contains_expected_measurements(self):
        model_results = evaluate_prediction_model(
            self.prediction_model,
            self.testing_hours,
            self.testing_marks,
        )

        self.assertIn("r_squared", model_results)
        self.assertIn("mean_absolute_error", model_results)
        self.assertIn("mean_squared_error", model_results)
        self.assertGreater(model_results["r_squared"], 0.90)

    def test_prediction_is_limited_to_valid_marks_range(self):
        predicted_marks = predict_student_marks(self.prediction_model, 24)

        self.assertGreaterEqual(predicted_marks, 0)
        self.assertLessEqual(predicted_marks, 100)


if __name__ == "__main__":
    unittest.main()
