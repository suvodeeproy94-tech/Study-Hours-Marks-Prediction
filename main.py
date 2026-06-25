"""
Study Hours vs Marks Prediction

This file trains a Linear Regression model that predicts a student's marks
from the number of hours studied.

Main responsibilities:
1. Create a small sample dataset.
2. Split the dataset into training and testing parts.
3. Train and evaluate the machine learning model.
4. Accept and validate study hours entered by the user.
5. Predict marks and display a graph.
"""

import math
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.model_selection import train_test_split


# These constants keep important values in one easy-to-find place.
MINIMUM_STUDY_HOURS = 0
MAXIMUM_STUDY_HOURS = 24
MINIMUM_MARKS = 0
MAXIMUM_MARKS = 100
TEST_DATA_SIZE = 0.25
RANDOM_STATE = 42
GRAPH_FILE = Path("images/study_hours_marks_graph.png")


def create_student_dataset():
    """
    Create and return the sample student dataset.

    The marks do not increase by exactly the same amount every time.
    This small variation makes the data more realistic than a perfect pattern.
    """

    student_data = {
        "Hours": [
            1.0,
            1.5,
            2.0,
            2.5,
            3.0,
            3.5,
            4.0,
            4.5,
            5.0,
            5.5,
            6.0,
            6.5,
            7.0,
            7.5,
            8.0,
            8.5,
            9.0,
            9.5,
            10.0,
            10.5,
        ],
        "Marks": [
            12,
            18,
            21,
            27,
            31,
            38,
            43,
            47,
            54,
            56,
            63,
            67,
            72,
            76,
            81,
            85,
            88,
            92,
            95,
            98,
        ],
    }

    return pd.DataFrame(student_data)


def prepare_training_and_testing_data(student_dataset):
    """
    Separate the dataset into input, output, training, and testing data.

    The model learns from the training data. The testing data is kept separate
    so that we can check the model using information it did not train on.
    """

    study_hours = student_dataset[["Hours"]]
    student_marks = student_dataset["Marks"]

    return train_test_split(
        study_hours,
        student_marks,
        test_size=TEST_DATA_SIZE,
        random_state=RANDOM_STATE,
    )


def train_prediction_model(training_hours, training_marks):
    """Create a Linear Regression model and train it."""

    prediction_model = LinearRegression()
    prediction_model.fit(training_hours, training_marks)
    return prediction_model


def evaluate_prediction_model(prediction_model, testing_hours, testing_marks):
    """
    Test the model and return three common regression measurements.

    R-squared shows how well the model follows the data pattern.
    Mean Absolute Error shows the average prediction difference in marks.
    Mean Squared Error gives more importance to larger prediction errors.
    """

    predicted_test_marks = prediction_model.predict(testing_hours)

    model_results = {
        "predicted_test_marks": predicted_test_marks,
        "r_squared": r2_score(testing_marks, predicted_test_marks),
        "mean_absolute_error": mean_absolute_error(
            testing_marks,
            predicted_test_marks,
        ),
        "mean_squared_error": mean_squared_error(
            testing_marks,
            predicted_test_marks,
        ),
    }

    return model_results


def validate_study_hours(entered_value):
    """
    Convert the entered value to a number and check its allowed range.

    A person cannot study for less than 0 or more than 24 hours in one day.
    A clear error is raised when the input is not valid.
    """

    if entered_value.strip() == "":
        raise ValueError("Study hours cannot be empty.")

    try:
        study_hours = float(entered_value)
    except ValueError as error:
        raise ValueError("Study hours must be a number, such as 5 or 5.5.") from error

    if not math.isfinite(study_hours):
        raise ValueError("Study hours must be a normal number.")

    if study_hours < MINIMUM_STUDY_HOURS:
        raise ValueError("Study hours cannot be negative.")

    if study_hours > MAXIMUM_STUDY_HOURS:
        raise ValueError("Study hours cannot be more than 24 hours in one day.")

    return study_hours


def ask_for_study_hours():
    """Keep asking until the user enters valid study hours."""

    while True:
        entered_value = input(
            f"\nEnter study hours ({MINIMUM_STUDY_HOURS}-{MAXIMUM_STUDY_HOURS}): "
        )

        try:
            return validate_study_hours(entered_value)
        except ValueError as error:
            print(f"Invalid input: {error}")


def predict_student_marks(prediction_model, study_hours):
    """
    Predict marks for the entered study hours.

    A DataFrame with the original 'Hours' column name is used here. This keeps
    the prediction input in the same format that was used during training.
    """

    prediction_input = pd.DataFrame({"Hours": [study_hours]})
    predicted_marks = prediction_model.predict(prediction_input)[0]

    # A student's marks should always stay between 0 and 100.
    limited_marks = max(MINIMUM_MARKS, min(predicted_marks, MAXIMUM_MARKS))
    return limited_marks


def display_model_equation(prediction_model):
    """Print the simple mathematical equation learned by the model."""

    coefficient = prediction_model.coef_[0]
    intercept = prediction_model.intercept_

    print("\nEquation learned by the model:")
    print(
        "Predicted Marks = "
        f"({coefficient:.2f} × Study Hours) + ({intercept:.2f})"
    )


def save_and_display_graph(student_dataset, prediction_model):
    """
    Create the study-hours graph, save it as an image, and display it.

    Blue points show the original data. The red line shows the predictions
    made by the trained Linear Regression model.
    """

    study_hours = student_dataset[["Hours"]]
    student_marks = student_dataset["Marks"]
    regression_line = prediction_model.predict(study_hours)

    GRAPH_FILE.parent.mkdir(parents=True, exist_ok=True)

    plt.figure(figsize=(9, 6))
    plt.scatter(
        student_dataset["Hours"],
        student_marks,
        color="royalblue",
        label="Original student data",
    )
    plt.plot(
        student_dataset["Hours"],
        regression_line,
        color="red",
        linewidth=2,
        label="Linear Regression line",
    )
    plt.title("Study Hours vs Marks Prediction")
    plt.xlabel("Study Hours")
    plt.ylabel("Marks")
    plt.xlim(left=0)
    plt.ylim(0, 105)
    plt.grid(alpha=0.3)
    plt.legend()
    plt.tight_layout()
    plt.savefig(GRAPH_FILE, dpi=150)

    print(f"\nGraph saved at: {GRAPH_FILE}")

    # The Agg backend is used during automated testing and cannot open a window.
    if "agg" not in plt.get_backend().lower():
        plt.show()

    plt.close()


def print_model_results(testing_marks, model_results):
    """Display actual values, predictions, and model measurements clearly."""

    comparison_table = pd.DataFrame(
        {
            "Actual Marks": testing_marks.to_numpy(),
            "Predicted Marks": model_results["predicted_test_marks"].round(2),
        }
    )

    print("\nTest data comparison:")
    print(comparison_table.to_string(index=False))

    print("\nModel evaluation:")
    print(f"R-squared score       : {model_results['r_squared']:.3f}")
    print(
        "Mean Absolute Error  : "
        f"{model_results['mean_absolute_error']:.2f} marks"
    )
    print(
        "Mean Squared Error   : "
        f"{model_results['mean_squared_error']:.2f}"
    )


def main():
    """Run the complete Study Hours vs Marks Prediction program."""

    print("=" * 55)
    print("STUDY HOURS VS MARKS PREDICTION")
    print("=" * 55)

    student_dataset = create_student_dataset()

    (
        training_hours,
        testing_hours,
        training_marks,
        testing_marks,
    ) = prepare_training_and_testing_data(student_dataset)

    prediction_model = train_prediction_model(training_hours, training_marks)
    print("\nModel trained successfully.")

    model_results = evaluate_prediction_model(
        prediction_model,
        testing_hours,
        testing_marks,
    )
    print_model_results(testing_marks, model_results)
    display_model_equation(prediction_model)

    entered_study_hours = ask_for_study_hours()
    predicted_marks = predict_student_marks(
        prediction_model,
        entered_study_hours,
    )

    print(
        f"\nPredicted marks for {entered_study_hours:g} study hours: "
        f"{predicted_marks:.2f} out of 100"
    )

    save_and_display_graph(student_dataset, prediction_model)


# This condition runs the program only when this file is opened directly.
# It also allows the functions to be imported safely by the test file.
if __name__ == "__main__":
    main()
