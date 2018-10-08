Attribute VB_Name = "Errors"
' Version 1.1.1

'@Folder("CodeBase.Constants.Errors")

Option Explicit

Public Const ERR_NUMBER_HEADERS_MISSING As Long = 513
Public Const ERR_MESSAGE_HEADERS_MISSING As String = "Supplied Header arguments not found within the headers of the provided Data."

Public Const ERR_NUMBER_FUNCTIONALITY_NOT_IMPLEMENTED As Long = 514
Public Const ERR_MESSAGE_FUNCTIONALITY_NOT_IMPLEMENTED As String = "This functionality has not yet been implemented. Please seek assistance before continuing."

Public Const ERR_NUMBER_COURSE_LIST_REQUIRED As Long = 515
Public Const ERR_MESSAGE_COURSE_LIST_REQUIRED As String = "Course list is required for this functionality. Please ensure a course list is provided, and try again."

Public Const ERR_NUMBER_COURSE_LIST_EMPTY As Long = 516
Public Const ERR_MESSAGE_COURSE_LIST_EMPTY As String = "A course list with at least one course must be provided in order to use this functionality. Please ensure that the course list is populated, and then try again."

Public Const ERR_NUMBER_INDEX_IS_ZERO As Long = 517
Public Const ERR_MESSAGE_INDEX_IS_ZERO As String = "The requested index is 0 and cannot be accessed."

Public Const ERR_NUMBER_RETRIEVER_NOT_SET As Long = 518
Public Const ERR_MESSAGE_RETRIEVER_NOT_SET As String = "This functionality requires a retriever to be set and loaded beforehand. Please ensure that the proper retriever is loaded."

Public Const ERR_NUMBER_FACTORY_TYPE_NOT_DEFINED As Long = 519
Public Const ERR_MESSAGE_FACTORY_TYPE_NOT_DEFINED As String = "Please ensure that the FactoryType provided is listed within the FactoryType Enum, and that functionality for the select FactoryType has been implemented."

Public Const ERR_NUMBER_RANGE_TYPE_NOT_DEFINED As Long = 520
Public Const ERR_MESSAGE_RANGE_TYPE_NOT_DEFINED As String = "Please ensure that the RangeType provided is listed within the RangeType Enum, and that functionality for the selected RangeType has been implemented."

Public Const ERR_NUMBER_ARRAY_EXPECTED As Long = 521
Public Const ERR_MESSAGE_ARRAY_EXPECTED As String = "The argument expects an array, but was not provided one. Please ensure that the argument provided is in the correct format."

Public Const ERR_NUMBER_WRONG_NUMBER_OF_ARRAY_DIMENTSIONS As Long = 522
Public Const ERR_MESSAGE_WRONG_NUMBER_OF_ARRAY_DIMENSIONS As String = "The provided array does not conform to the expected size. Please check the provided arguments."

Public Const ERR_NUMBER_BASE_DIRECTORY_MISSING As Long = 523
Public Const ERR_MESSAGE_BASE_DIRECTORY_MISSING As String = "The provided directory-path is missing. Please ensure that the file-path is correct, and then try again."

Public Const ERR_NUMBER_NO_REPORT_WEEKS_REQUESTED As Long = 524
Public Const ERR_MESSAGE_NO_REPORT_WEEKS_REQUESTED As String = "Please check to make sure that you have properly filled in the weeks you would like included."

Public Const ERR_NUMBER_TERM_WEEK_GREATER_THAN_MAXIMUM As Long = 525
Public Const ERR_MESSAGE_TERM_WEEK_GREATER_THAN_MAXIMUM As String = "The provided term week number is too large. Please ensure the entered value is correct."
