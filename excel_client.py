import logging
import pandas as pd

import constants as c

class ExcelCient():
    def __init__(self):
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def to_df(file_name: str, sheet_name: str, header: bool = True) -> pd.DataFrame:
        """
        Read Excel file into a Pandas DataFrame.
        Args:
            file (str): Path to the Excel file.
            sheet (str): Name of the sheet to read from.
            header (bool, optional): Whether to use the first row as column headers
                Defaults to True.
        Returns:
            pd.DataFrame: The read Excel data in a Pandas DataFrame format.
        Raises:
            FileNotFoundError: If the specified file does not exist.
            Exception: Any other exception raised during file readout or parsing.
        """
        try:
            # Attempt to read the Excel file, handling potential exceptions
            if header:
                df = pd.read_excel(file_name, sheet_name=sheet_name, header=0)
            else:
                df = pd.read_excel(file_name, sheet_name=sheet_name)
            
            # Return the DataFrame with the Excel data
            return df
        
        except FileNotFoundError as e:
            # Raise a specific error message for non-existent files
            self.logger.error(f"File '{file_name}' not found.")
            raise FileNotFoundError(f"File '{file_name}' not found.") from e
        
        except Exception as e:
            # Raise other exceptions that occurred during file readout or parsing
            self.logger.error(f"Exception raised during file readout: {e}")
            raise Exception(f"Exception raised during file readout: {e}") from e

    def validate_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Validate the input for missing values and perform datatype conversion.
        Args:
            df (pd.DataFrame): The input DataFrame to be validated.
        Returns:
            pd.DataFrame: The validated and converted DataFrame.
        Raises:
            exception: If any error occurs during datatype conversion.
        """
        # Check for missing values in the input dataset
        missing_values = df.isna().any(axis=1)

        # If there are any missing values, raise a ValueError with detailed message and affected rows
        if missing_values.any():
            raise ValueError(
                f"Missing values found in the input dataset:\n{df.loc[missing_values]}\n"
            )
        else:
            self.logger.debug("Dataset loaded, no apparent missing values.")
    
        # Attempt to convert DataFrame datatypes according to the FEATURES constant
        try:
            df = df.astype(c.FEATURES)
            return df
        except Exception as e:
            self.logger.error(f"Exception raised during datatype conversion: {e}")
            raise ValueError(f"Exception raised during datatype conversion: {e}") from e

    def format_df(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Formats a DataFrame by selecting features, filtering out empty rows,
        formatting feature datatypes and validating the dataset.
        Args:
            df: A Pandas DataFrame to be formatted.
        Returns:
            The formatted DataFrame.
        """
        # Select only the desired features from the DataFrame
        # This filtering method avoids unnecessary computation and memory usage
        df = df.loc[:, c.FEATURES.keys()]

        # Filter out empty rows based on 'Name' column
        # It's more efficient to filter before dropping NaN values
        filtered_df = df[df['Name'].notna()].reset_index(drop=True)

        # Format feature datatypes and validate dataset
        # It's a good practice to separate formatting from validation
        formatted_df = self.validate_data(filtered_df)

        return formatted_df

    def load_excel(self, file_name: str, sheet_name: str, header: bool = True) -> pd.DataFrame:
        """
        Read in Excel file, format and validate data.
        Args:
            file (str): Path to the Excel file.
            sheet (str): Name of the sheet to read from.
            header (bool, optional): Whether to use the first row as column headers
                Defaults to True.
        Returns:
            pd.DataFrame: The formatted and validated DataFrame.
        Raises:
            FileNotFoundError: If the specified file does not exist.
            Exception: Any other exception raised during file readout or parsing.
        """
        # Read in Excel file
        logging.info("Reading in Excel file")
        df = excel_to_df(file_name=file_name,
                        sheet_name=sheet_name,
                        header=header)

        self.logger.info("Reading in Excel file")
        df = self.excel_to_df(file_name=file_name, sheet_name=sheet_name, header=header)
        # Formatting - feature selectioncleaning and validation
        logging.info("Formatting and validating data")
        df = format_df(df)
        self.logger.info("Formatting and validating data")
        df = self.format_df(df)

        return df

# Example usage:
excel_client = ExcelClient()
df = excel_client.load_excel(file_name='example.xlsx', sheet_name='Sheet1')