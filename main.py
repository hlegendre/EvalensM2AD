import pandas as pd
import os
import sacred

ex = sacred.Experiment("SummarizeResult", save_git_info=False)


@ex.config
def config():
    input_folder = "input"
    output_folder = "output"
    output_prefix = "[PROCESSED] "


@ex.capture
def process_file(input_name, output_name, output_prefix):
    # ------ Reading the file ------ #
    evalens_file = pd.read_excel(input_name, header=None)

    # ------ Determining the questions ------ #
    # Distinguishing between the open questions and the normal questions
    questions_indices = evalens_file[evalens_file[0].str.startswith("Résumé pour Q").fillna(False)].index
    questions_libre_indices = evalens_file[evalens_file[0].str.startswith("Identifiant (ID)").fillna(False)].index
    questions_libre_indices = [max([x for x in questions_indices if x < i]) for i in questions_libre_indices]
    questions_indices = [x for x in questions_indices if x not in questions_libre_indices]


    # ------ Creating the result frame ------ #
    result_dataframe = pd.DataFrame(
        columns=["Intitulé Question", "Très Satisfaits", "Satisfaits", "Scores"],
    )
    # save the colors separately
    score_colors = {}

    for i, index in enumerate(questions_indices):
        # identifying the question
        question_id = evalens_file.iloc[index, 0].split(' ')[-1]
        question = evalens_file.iloc[index + 1, 0]
        result_dataframe.loc[question_id, "Intitulé Question"] = question

        # very satisfied
        n_tres_satisfait = evalens_file.iloc[index+3, 1]
        n_satisfait = evalens_file.iloc[index+4, 1]
        n_autre = evalens_file.iloc[index+5, 1]\
             + evalens_file.iloc[index+6, 1]\

        tres_satisfait = n_tres_satisfait / (n_tres_satisfait + n_satisfait + n_autre)
        result_dataframe.loc[question_id, "Très Satisfaits"] = tres_satisfait

        # satisfied
        satisfait = n_satisfait / (n_tres_satisfait + n_satisfait + n_autre)
        result_dataframe.loc[question_id, "Satisfaits"] = satisfait

        # sum of both
        score = tres_satisfait + satisfait
        result_dataframe.loc[question_id, "Scores"] = score

        # determining the color
        if 0.8 <= score and tres_satisfait >= satisfait:
            color = "#0000FF"
        elif 0.8 <= score:
            color = "#00B050"
        elif 0.7 <= score < 0.8:
            color = "#FFFF00"
        elif 0.5 <= score < 0.7:
            color = "#FFBF00"
        else:
            color = "#FF0000"
        score_colors[question_id] = color

    # ------ Creating the output excel file ------ #

    writer = pd.ExcelWriter(
        os.path.join(os.path.dirname(output_name), output_prefix + os.path.basename(output_name)[:-4] + ".xlsx"),
        engine='xlsxwriter'
    )

    workbook = writer.book

    # Because of the formating, we have to write the synthesis sheet manually
    start_col = 1
    start_row = 15
    result_dataframe.to_excel(writer, sheet_name="Synthèse", startrow=start_row, startcol=start_col)
    worksheet = writer.sheets['Synthèse']

    header_format = workbook.add_format({
        'bold': False,
        'text_wrap': False,
        'valign': 'top',
        'bg_color': '#D0CECE',
        'border': 1})

    # Write the column titles
    worksheet.write(start_row, start_col, "", header_format)
    for col_num, value in enumerate(result_dataframe.columns.values):
        worksheet.write(start_row, start_col + col_num + 1, value, header_format)

    for i, question in enumerate(score_colors.keys()):
        worksheet.write(
            start_row + i + 1,
            start_col,
            question,
            workbook.add_format({'border': 1})
        )
        worksheet.write(
            start_row + i + 1,
            start_col + 1,
            result_dataframe.loc[question, "Intitulé Question"],
            workbook.add_format({'border': 1})
        )
        worksheet.write(
            start_row + i + 1,
            start_col + 2,
            result_dataframe.loc[question, "Très Satisfaits"],
            workbook.add_format({'border': 1, 'num_format': '0%'})
        )
        worksheet.write(
            start_row + i + 1,
            start_col + 3,
            result_dataframe.loc[question, "Satisfaits"],
            workbook.add_format({'border': 1, 'num_format': '0%'})
        )
        worksheet.write(
            start_row + i + 1,
            start_col + 4,
            result_dataframe.loc[question, "Scores"],
            workbook.add_format({'border': 1, 'num_format': '0%', 'bg_color': score_colors[question]})
        )

    # ------ Writing the legend ------ #
    worksheet.write(
            1,
            1,
            "",
            workbook.add_format({'bg_color': "#0000FF"})
        )
    worksheet.write(
        1,
        2,
        "Plus de 80% de satisfaits ou très satisfaits ET autant ou plus de très satisfaits que satisfaits",
    )

    worksheet.write(
            2,
            1,
            "",
            workbook.add_format({'bg_color': "#00B050"})
        )
    worksheet.write(
        2,
        2,
        "Plus de 80% de satisfaits ou très satisfaits ET moins de très satisfaits que satisfaits",
    )

    worksheet.write(
            3,
            1,
            "",
            workbook.add_format({'bg_color': "#FFFF00"})
        )
    worksheet.write(
        3,
        2,
        "Entre 70% inclus et 80% exclus de satisfaits ou très satisfaits",
    )

    worksheet.write(
            4,
            1,
            "",
            workbook.add_format({'bg_color': "#FFBF00"})
        )
    worksheet.write(
        4,
        2,
        "Entre 50% inclus et 70% exclus de satisfaits ou très satisfaits",
    )

    worksheet.write(
            5,
            1,
            "",
            workbook.add_format({'bg_color': "#FF0000"})
        )
    worksheet.write(
        5 ,
        2,
        "Moins de 50% exclus de satisfaits ou très satisfaits",
    )

    # ------ writing the participation summary ------ #
    worksheet.write(
        10 ,
        2,
        "Nombre de réponses",
    )
    worksheet.write(
        10 ,
        3,
        evalens_file.iloc[0, 1],
    )
    worksheet.write(
        11 ,
        2,
        "Pourcentage",
    )
    worksheet.write(
        11 ,
        3,
        evalens_file.iloc[2, 1],
        workbook.add_format({'num_format': '0%'})
    )

    # Copying the evalens file in another sheet
    evalens_file.to_excel(writer, sheet_name="Détails et commentaires", index=False, header=False)
    writer.save()


@ex.automain
def main(input_folder, output_folder):
    for file_name in os.listdir(input_folder):
        if file_name!=".gitignore":
            process_file(
                input_name=os.path.join(input_folder, file_name),
                output_name=os.path.join(output_folder, file_name)
            )
    print("All the files in {} have been processed".format(input_folder))
    print("Check the processed file in {}".format(output_folder))
