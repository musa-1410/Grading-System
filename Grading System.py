import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
import numpy as np
from scipy.stats import norm
import matplotlib.pyplot as plt


class GradingSystemApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Grading System")

        self.data = None
        self.weightages = None
        self.total_marks = None
        self.default_stdev_adjustment = 1.0

        # Updated pre-adjustment boundaries for absolute grading
        self.default_boundaries = {
            90: "A",
            85: "A-",
            80: "B+",
            75: "B",
            70: "B-",
            65: "C+",
            60: "C",
            55: "C-",
            50: "D",
            0: "F"
        }

        # Input Section
        tk.Label(root, text="Upload Excel File").grid(row=0, column=0, padx=10, pady=10)
        self.upload_button = tk.Button(root, text="Upload", command=self.upload_file)
        self.upload_button.grid(row=0, column=1, padx=10, pady=10)

        # Grading Scheme
        tk.Label(root, text="Select Grading Scheme").grid(row=1, column=0, padx=10, pady=10)
        self.grading_scheme = tk.StringVar(value="absolute")
        tk.Radiobutton(root, text="Absolute", variable=self.grading_scheme, value="absolute").grid(row=1, column=1)
        tk.Radiobutton(root, text="Relative", variable=self.grading_scheme, value="relative").grid(row=1, column=2)

        # Process and Save Buttons
        self.process_button = tk.Button(root, text="Process Grades", command=self.process_grades)
        self.process_button.grid(row=2, column=0, pady=20)
        self.save_button = tk.Button(root, text="Save Results", command=self.save_results)
        self.save_button.grid(row=2, column=1, pady=20)

    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            try:
                self.data = pd.read_excel(file_path, engine="openpyxl")
                required_columns = ['Quizzes', 'Assignments', 'Midterm', 'Finals', 'Project']
                if not all(col in self.data.columns for col in required_columns):
                    messagebox.showerror(
                        "Error",
                        "Excel file must contain required columns: Quizzes, Assignments, Midterm, Finals, Project.",
                    )
                    self.data = None
                else:
                    messagebox.showinfo("Success", "File uploaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {e}")

    def input_weightages(self):
        weightages = {}
        try:
            weightages["quizzes"] = float(simpledialog.askstring("Input", "Enter weightage for Quizzes (%): "))
            weightages["assignments"] = float(simpledialog.askstring("Input", "Enter weightage for Assignments (%): "))
            weightages["midterm"] = float(simpledialog.askstring("Input", "Enter weightage for Midterm (%): "))
            weightages["final"] = float(simpledialog.askstring("Input", "Enter weightage for Final (%): "))
            weightages["project"] = float(simpledialog.askstring("Input", "Enter weightage for Project (%): "))

            if sum(weightages.values()) != 100:
                raise ValueError("Total weightages must equal 100%.")
        except ValueError as e:
            messagebox.showerror("Input Error", str(e))
            return None
        return weightages

    def input_total_marks(self):
        total_marks = {}
        try:
            total_marks["quizzes"] = float(simpledialog.askstring("Input", "Enter total marks for Quizzes: "))
            total_marks["assignments"] = float(simpledialog.askstring("Input", "Enter total marks for Assignments: "))
            total_marks["midterm"] = float(simpledialog.askstring("Input", "Enter total marks for Midterm: "))
            total_marks["final"] = float(simpledialog.askstring("Input", "Enter total marks for Final: "))
            total_marks["project"] = float(simpledialog.askstring("Input", "Enter total marks for Project: "))
        except ValueError as e:
            messagebox.showerror("Input Error ", str(e))
            return None
        return total_marks

    def calculate_final_scores(self):
        try:
            self.data['Final Score'] = (
                (self.data['Quizzes'] / self.total_marks['quizzes']) * self.weightages['quizzes'] +
                (self.data['Assignments'] / self.total_marks['assignments']) * self.weightages['assignments'] +
                (self.data['Midterm'] / self.total_marks['midterm']) * self.weightages['midterm'] +
                (self.data['Finals'] / self.total_marks['final']) * self.weightages['final'] +
                (self.data['Project'] / self.total_marks['project']) * self.weightages['project']
            )
        except KeyError as e:
            messagebox.showerror("Error", f"Missing column in input data: {e}")

    def apply_absolute_grading(self, scores, boundaries):
        grades = []
        for score in scores:
            for boundary in sorted(boundaries.keys(), reverse=True):
                if score >= boundary:
                    grades.append(boundaries[boundary])
                    break
        return grades

    def apply_relative_grading(self, scores, stdev_adjustment):
        mean = np.mean(scores)
        std_dev = np.std(scores)
        grades = []

        for score in scores:
            if score >= mean + 2 * std_dev:
                grades.append("A")
            elif mean + (3 / 2) * std_dev <= score < mean + 2 * std_dev:
                grades.append("A-")
            elif mean + std_dev <= score < mean + (3 / 2) * std_dev:
                grades.append("B+")
            elif mean + (1 / 2) * std_dev <= score < mean + std_dev:
                grades.append("B")
            elif mean - (1 / 2) * std_dev <= score < mean + (1 / 2) * std_dev:
                grades.append("B-")
            elif mean - std_dev <= score < mean - (1 / 2) * std_dev:
                grades.append("C+")
            elif mean - (4 / 3) * std_dev <= score < mean - std_dev:
                grades.append("C")
            elif mean - (5 / 3) * std_dev <= score < mean - (4 / 3) * std_dev:
                grades.append("C-")
            elif mean - 2 * std_dev <= score < mean - (5 / 3) * std_dev:
                grades.append("D")
            else:
                grades.append("F")
        return grades

    def apply_post_adjustment_grading(self, scores, stdev_multiplier):
        mean = np.mean(scores)
        std_dev = np.std(scores)
        grades = []

        grade_boundaries = [
            ("A", mean + 1.5 * std_dev),
            ("A-", mean + (1.5 * std_dev - stdev_multiplier * std_dev)),
            ("B+", mean + (1.5 * std_dev - 2 * stdev_multiplier * std_dev)),
            ("B", mean + (1.5 * std_dev - 3 * stdev_multiplier * std_dev)),
            ("B-", mean + (1.5 * std_dev - 4 * stdev_multiplier * std_dev)),
            ("C+", mean + (1.5 * std_dev - 5 * stdev_multiplier * std_dev)),
            ("C", mean + (1.5 * std_dev - 6 * stdev_multiplier * std_dev)),
            ("C-", mean + (1.5 * std_dev - 7 * stdev_multiplier * std_dev)),
            ("D", mean + (1.5 * std_dev - 8 * stdev_multiplier * std_dev)),
            ("F", float("-inf"))  # F applies to all scores below the last boundary
        ]

        for score in scores:
            for grade, boundary in grade_boundaries:
                if score >= boundary:
                    grades.append(grade)
                    break

        return grades

    def visualize_grades(self, pre_grades, post_grades, title):
        # Count frequencies of each grade
        pre_grade_counts = pd.Series(pre_grades).value_counts().sort_index()
        post_grade_counts = pd.Series(post_grades).value_counts().sort_index()

        # Plot grade distribution side by side
        fig, axes = plt.subplots(1, 2, figsize=(12, 6))
        pre_grade_counts.plot(kind="bar", alpha=0.7, color="skyblue", ax=axes[0])
        post_grade_counts.plot(kind="bar", alpha=0.7, color="lightgreen", ax=axes[1])

        axes[0].set_title("Pre-Adjustment Grades")
        axes[0].set_xlabel("Grades")
        axes[0].set_ylabel("Frequency")
        axes[1].set_title("Post-Adjustment Grades")
        axes[1].set_xlabel("Grades")
        axes[1].set_ylabel("Frequency")

        plt.suptitle(title)
        plt.tight_layout()
        plt.show()

    def visualize_bell_curve(self, scores, title):
        mean = np.mean(scores)
        std_dev = np.std(scores)

        x = np.linspace(mean - 3 * std_dev, mean + 3 * std_dev, 1000)
        y = norm.pdf(x, mean, std_dev)

        plt.figure(figsize=(10, 6))
        plt.plot(x, y, label="Bell Curve", color="red", linewidth=2)
        plt.xlabel("Grades")
        plt.ylabel("Probability Density")
        plt.title(f"{title}\nMean: {mean:.2f}, Std Dev: {std_dev:.2f}")
        plt.axvline(mean, color="blue", linestyle="--", label="Mean")
        plt.axvline(mean + std_dev, color="green", linestyle="--", label="Mean + 1 SD")
        plt.axvline(mean - std_dev, color="green", linestyle="--", label="Mean - 1 SD")
        plt.axvline(mean + 2 * std_dev, color="orange", linestyle="--", label="Mean + 2 SD")
        plt.axvline(mean - 2 * std_dev, color="orange", linestyle="--", label="Mean - 2 SD")
        plt.axvline(mean + 3 * std_dev, color="purple", linestyle="--", label="Mean + 3 SD")
        plt.axvline(mean - 3 * std_dev, color="purple", linestyle="--", label="Mean - 3 SD")
        plt.legend()
        plt.show()

    def input_absolute_boundaries(self):
        boundaries = {}
        try:
            prev_boundary = float('inf')  # Initialize with a large value to ensure the first input is valid
            while True:
                grade = simpledialog.askstring("Input", "Enter grade label (e.g., A, B, etc.): ")
                if not grade:
                    break
                boundary_input = simpledialog.askstring("Input", f"Enter lower boundary for grade {grade}: ")
                if not boundary_input:
                    break
                try:
                    boundary = float(boundary_input)
                    if boundary >= prev_boundary:
                        messagebox.showerror("Input Error", f"The lower boundary for {grade} must be less than the previous boundary.")
                        continue  # Skip this iteration and ask for the boundary again
                    boundaries[boundary] = grade
                    prev_boundary = boundary  # Update prev_boundary to the current one
                except ValueError:
                    messagebox.showerror("Input Error", "Invalid input for boundary. Please enter a numeric value.")
                    continue  # Skip to the next iteration if invalid input is provided

                more = messagebox.askyesno("Continue", "Do you want to add another grade boundary?")
                if not more:
                    break

            # Sort boundaries in descending order of the numerical values
            boundaries = dict(sorted(boundaries.items(), reverse=True))
        except ValueError as e:
            messagebox.showerror("Input Error", "Invalid input for grade boundaries.")
            return None
        return boundaries


    def process_grades(self):
        if self.data is None:
            messagebox.showerror("Error", "Please upload a valid grades file.")
            return

        # Get weightages for the components
        self.weightages = self.input_weightages()
        if not self.weightages:
            return

        # Get total marks for the components
        self.total_marks = self.input_total_marks()
        if not self.total_marks:
            return

        # Calculate the final scores
        self.calculate_final_scores()

        scheme = self.grading_scheme.get()
        if scheme == "absolute":
            # Pre-adjustment grades based on the default boundaries
            pre_grades = self.apply_absolute_grading(self.data['Final Score'], self.default_boundaries)
            self.visualize_grades(pre_grades, pre_grades, "Pre-Adjustment Absolute Grading")

            # Ask the user if they want to adjust the grade boundaries
            adjust = messagebox.askyesno("Adjust Grades", "Do you want to adjust the grade boundaries?")
            if adjust:
                # Input custom boundaries for post-adjustment
                post_boundaries = self.input_absolute_boundaries()
                if post_boundaries:
                    # Post-adjustment grades based on user-defined boundaries
                    post_grades = self.apply_absolute_grading(self.data['Final Score'], post_boundaries)
                    self.visualize_grades(pre_grades, post_grades, "Post-Adjustment Absolute Grading")
                    self.data['Grade'] = post_grades  # Assign grades to the dataset
                    self.show_summary_statistics(pre_grades, post_grades)
                else:
                    messagebox.showerror("Error", "Invalid grade boundaries entered.")
                    return
            else:
                # Assign pre-adjustment grades if no adjustment is needed
                self.data['Grade'] = pre_grades

        elif scheme == "relative":
            # Pre-adjustment grades based on the default standard deviation
            pre_grades = self.apply_relative_grading(self.data['Final Score'], self.default_stdev_adjustment)
            self.visualize_grades(pre_grades, pre_grades, "Pre-Adjustment Relative Gr ading")
            self.visualize_bell_curve(self.data['Final Score'], "Pre-Adjustment Relative Grading Bell Curve")

            # Ask the user if they want to adjust the grading using a custom standard deviation multiplier
            adjust = messagebox.askyesno("Adjust Grades", "Do you want to adjust the grading using a custom standard deviation multiplier?")
            if adjust:
                stdev_multiplier = float(simpledialog.askstring("Input", "Enter standard deviation multiplier for post-adjustment (default is 1.0): "))
                post_grades = self.apply_post_adjustment_grading(self.data['Final Score'], stdev_multiplier)
                self.visualize_grades(pre_grades, post_grades, "Post-Adjustment Relative Grading")
                self.visualize_bell_curve(self.data['Final Score'], "Post-Adjustment Relative Grading Bell Curve")
                self.data['Grade'] = post_grades  # Assign grades to the dataset
                self.show_summary_statistics(pre_grades, post_grades)
            else:
                # Assign pre-adjustment grades if no adjustment is needed
                self.data['Grade'] = pre_grades

        # Inform the user that processing is complete
        messagebox.showinfo("Success", "Grades processed successfully!")

    def show_summary_statistics(self, pre_grades, post_grades):
        pre_grade_counts = pd.Series(pre_grades).value_counts()
        post_grade_counts = pd.Series(post_grades).value_counts()

        summary = "Summary Statistics:\n"
        summary += f"Total Students: {len(self.data)}\n"
        summary += "Grade Changes:\n"

        for grade in pre_grade_counts.index:
            pre_count = pre_grade_counts.get(grade, 0)
            post_count = post_grade_counts.get(grade, 0)
            if pre_count != post_count:
                summary += f"{grade}: {pre_count} -> {post_count}\n"

        messagebox.showinfo("Summary Statistics", summary)

    def save_results(self):
        if self.data is None or 'Grade' not in self.data.columns:
            messagebox.showerror("Error", "No data to save or grades not assigned.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", ".csv"), ("Excel Files", ".xlsx")])
        if file_path:
            try:
                # Save all input data along with the assigned grades
                output_data = self.data.copy()

                # Ensure only necessary columns are included
                if 'Final Score' not in output_data.columns:
                    output_data['Final Score'] = self.data['Final Score']
                if 'Grade' not in output_data.columns:
                    output_data['Grade'] = self.data['Grade']

                if file_path.endswith('.csv'):
                    output_data.to_csv(file_path, index=False)
                elif file_path.endswith('.xlsx'):
                    output_data.to_excel(file_path, index=False, engine='openpyxl')

                messagebox.showinfo("Success", f"Results saved to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = GradingSystemApp(root)
    root.mainloop()
