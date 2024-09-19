""" This module provides the ExcelClient class to handle Excel file data """

import logging
import pandas as pd

import excel_client.constants as c

class ExcelClient():
    """
    Class designed to load, format, validate and export data contained in Excel files
    """
    def __init__(self,
                 file_path: str):
        self.logger = logging.getLogger(__name__)
        self.file_path = file_path
        self.dataframe = pd.DataFrame()


    def _to_df(self,
              sheet_name: str,
              header: bool = True) -> None or pd.DataFrame:
        """
        Read Excel file into a Pandas DataFrame.
        Args:
            file (str): Path to the Excel file.
            sheet (str): Name of the sheet to read from.
            header (bool, optional): Whether to use the first row as column headers
                Defaults to True.
        Returns:
            None or pd.DataFrame
        Raises:
            FileNotFoundError: If the specified file does not exist.
            Exception: Any other exception raised during file readout or parsing.
        """
        try:
            # Attempt to read the Excel file, handling potential exceptions
            if header:
                self.dataframe = pd.read_excel(
                    self.file_path, sheet_name=sheet_name, header=0
                )
            else:
                self.dataframe = pd.read_excel(
                    self.file_path, sheet_name=sheet_name
                )

        except FileNotFoundError as e:
            # Raise a specific error message for non-existent files
            self.logger.error("File '%s' not found.", self.file_path)
            raise FileNotFoundError(f"File '{self.file_path}' not found.") from e

        except Exception as e:
            # Raise other exceptions that occurred during file readout or parsing
            self.logger.error("Exception raised during file readout: %s", e)
            raise ValueError(f"Exception raised during file readout: {e}") from e

    def _validate_data(self) -> None:
        """
        Validate the input for missing values and perform datatype conversion.
        Raises:
            exception: If any error occurs during datatype conversion.
        """
        # Check for missing values in the input dataset
        missing_values = self.dataframe.isna().any(axis=1)

        # If there are any missing values, raise a ValueError with detailed message
        # and affected rows
        if missing_values.any():
            raise ValueError(
                f"Missing values found in the input dataset: \
                    \n{self.dataframe.loc[missing_values]}\n"
            )
        else:
            self.logger.debug("Dataset loaded, no apparent missing values.")

        # Attempt to convert DataFrame datatypes according to the FEATURES constant
        try:
            self.dataframe = self.dataframe.astype(c.FEATURES)
        except Exception as e:
            self.logger.error("Exception raised during datatype conversion: %s", e)
            raise ValueError(f"Exception raised during datatype conversion: {e}") from e


    def _format_df(self, validate: bool = True) -> None:
        """
        Formats a DataFrame by selecting features, filtering out empty rows,
        formatting feature datatypes and validating the dataset.
        Args:
            validate: Whether to perform validation after formatting
        """
        # Select only the desired features from the DataFrame
        # This filtering method avoids unnecessary computation and memory usage
        self.dataframe = self.dataframe.loc[:, c.FEATURES.keys()]

        # Filter out empty rows based on 'Name' column
        # It's more efficient to filter before dropping NaN values
        self.dataframe = self.dataframe[self.dataframe['Name'].notna()].reset_index(drop=True)

        # Format feature datatypes and validate dataset
        # It's a good practice to separate formatting from validation
        if validate:
            self._validate_data()


    def load_excel(self,
                   sheet_name: str,
                   header: bool = True,
                   return_output: bool = False) -> None or pd.DataFrame:
        """
        Read in Excel file, format and validate data.
        Args:
            file (str): Path to the Excel file.
            sheet (str): Name of the sheet to read from.
            header (bool, optional): Whether to use the first row as column headers
                Defaults to True.
            return_output (bool): Whether to return the formatted DataFrame
        Returns:
            pd.DataFrame: The formatted and validated DataFrame.
        Raises:
            FileNotFoundError: If the specified file does not exist.
            Exception: Any other exception raised during file readout or parsing.
        """
        # Read in Excel file
        self.logger.info("Reading in Excel file")
        self._to_df(sheet_name=sheet_name, header=header)
        # Formatting - feature selectioncleaning and validation
        self.logger.info("Formatting and validating data")
        self._format_df(validate=True)

        returned_object = self.dataframe if return_output else None
        return returned_object
