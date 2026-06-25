"""
Research Evaluation for the Study Hours vs Marks Project

This file reproduces the extra model checks reported in the IEEE research paper.
It compares Linear Regression with a simple mean baseline and then performs
five-fold cross-validation.
"""

import numpy as np
from sklearn.dummy import DummyRegressor
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.model_selection import KFold, cross_validate, train_test_split

from main import RANDOM_STATE, TEST_DATA_SIZE, create_student_dataset


def calculate_regression_results(model, training_hours, testing_hours,
                                 training_marks, testing_marks):
    """Train one model and return its three regression measurements."""

    model.fit(training_hours, training_marks)
    predicted_marks = model.predict(testing_hours)

    results = {
        "r_squared": r2_score(testing_marks, predicted_marks),
        "mean_absolute_error": mean_absolute_error(
            testing_marks,
            predicted_marks,
        ),
        "root_mean_squared_error": np.sqrt(
            mean_squared_error(testing_marks, predicted_marks)
        ),
    }

    return results


def run_fixed_split_comparison(study_hours, student_marks):
    """Compare Linear Regression with a simple mean-value baseline."""

    (
        training_hours,
        testing_hours,
        training_marks,
        testing_marks,
    ) = train_test_split(
        study_hours,
        student_marks,
        test_size=TEST_DATA_SIZE,
        random_state=RANDOM_STATE,
    )

    linear_results = calculate_regression_results(
        LinearRegression(),
        training_hours,
        testing_hours,
        training_marks,
        testing_marks,
    )

    baseline_results = calculate_regression_results(
        DummyRegressor(strategy="mean"),
        training_hours,
        testing_hours,
        training_marks,
        testing_marks,
    )

    return linear_results, baseline_results


def run_five_fold_cross_validation(study_hours, student_marks):
    """Run five train-test rounds and return the mean and spread of each result."""

    five_folds = KFold(
        n_splits=5,
        shuffle=True,
        random_state=RANDOM_STATE,
    )

    fold_results = cross_validate(
        LinearRegression(),
        study_hours,
        student_marks,
        cv=five_folds,
        scoring={
            "r_squared": "r2",
            "mean_absolute_error": "neg_mean_absolute_error",
            "mean_squared_error": "neg_mean_squared_error",
        },
    )

    r_squared_values = fold_results["test_r_squared"]
    mean_absolute_error_values = -fold_results["test_mean_absolute_error"]
    root_mean_squared_error_values = np.sqrt(
        -fold_results["test_mean_squared_error"]
    )

    return {
        "r_squared_mean": np.mean(r_squared_values),
        "r_squared_standard_deviation": np.std(r_squared_values),
        "mean_absolute_error_mean": np.mean(mean_absolute_error_values),
        "mean_absolute_error_standard_deviation": np.std(
            mean_absolute_error_values
        ),
        "root_mean_squared_error_mean": np.mean(
            root_mean_squared_error_values
        ),
        "root_mean_squared_error_standard_deviation": np.std(
            root_mean_squared_error_values
        ),
    }


def print_model_results(model_name, results):
    """Display one model's results in a clear format."""

    print(f"\n{model_name}")
    print("-" * len(model_name))
    print(f"R-squared : {results['r_squared']:.3f}")
    print(f"MAE       : {results['mean_absolute_error']:.2f} marks")
    print(f"RMSE      : {results['root_mean_squared_error']:.2f} marks")


def main():
    """Run and display all research evaluation checks."""

    student_dataset = create_student_dataset()
    study_hours = student_dataset[["Hours"]]
    student_marks = student_dataset["Marks"]

    linear_results, baseline_results = run_fixed_split_comparison(
        study_hours,
        student_marks,
    )
    cross_validation_results = run_five_fold_cross_validation(
        study_hours,
        student_marks,
    )

    print("=" * 58)
    print("RESEARCH EVALUATION RESULTS")
    print("=" * 58)

    print_model_results("Linear Regression", linear_results)
    print_model_results("Mean Baseline", baseline_results)

    print("\nFive-Fold Cross-Validation")
    print("--------------------------")
    print(
        "R-squared : "
        f"{cross_validation_results['r_squared_mean']:.3f} "
        f"+/- {cross_validation_results['r_squared_standard_deviation']:.3f}"
    )
    print(
        "MAE       : "
        f"{cross_validation_results['mean_absolute_error_mean']:.2f} "
        f"+/- "
        f"{cross_validation_results['mean_absolute_error_standard_deviation']:.2f} "
        "marks"
    )
    print(
        "RMSE      : "
        f"{cross_validation_results['root_mean_squared_error_mean']:.2f} "
        f"+/- "
        f"{cross_validation_results['root_mean_squared_error_standard_deviation']:.2f} "
        "marks"
    )


if __name__ == "__main__":
    main()
