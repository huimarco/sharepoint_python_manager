import pandas as pd

def print_progress(items):
    '''
    Print the number of line items read so far.
    '''
    print('Items read: {0}'.format(len(items)))

def get_column_names(sp_list):
    '''
    Retrieve the actual column names (internal names) from a SharePoint list.
    '''
    # Fetch all fields for the SharePoint list
    fields = sp_list.fields.get().execute_query()
    # Extract the internal names of the fields
    column_names = {field.internal_name: field.title for field in fields}
    return column_names

def remove_cpu_columns(df):
    '''
    Drop irrelevant columns from output dataframe.
    '''
    # List of columns to drop
    columns_to_drop = [
        'ParentList', 'FileSystemObjectType', 'Id', 'ServerRedirectedEmbedUri', 
        'ServerRedirectedEmbedUrl', 'Content Type ID', 'OData__ColorTag', 
        'Compliance Asset Id', 'ID', 'Modified', 'Created', 'AuthorId', 
        'EditorId', 'OData__UIVersionString', 'Attachments', 'GUID'
    ]
    # Drop the columns
    return df.drop(columns=[col for col in columns_to_drop if col in df.columns])


def list_to_df(sp_list):
    '''
    Return all items in a Sharepoint list as a pandas dataframe.
    Maybe drop unecessary columns and data before converting to pandas dataframe...
    '''
    # Retrieve items from the list in pages of 500, and display progress using 'print_progress' callback.
    paged_items = sp_list.items.paged(500, page_loaded=print_progress).get().execute_query()
    # Retrieve the column names (mapping internal names to display names)
    column_names = get_column_names(sp_list)
    # Initialize an empty list to store the data
    data = []  
    for index, item in enumerate(paged_items):
        # Map the internal names to the actual column names
        row = {column_names.get(field, field): value for field, value in item.properties.items()}
        # Record data
        data.append(row)
    # Save data as pandas dataframe
    df = pd.DataFrame(data)
    # Drop unecessary columns and return
    return remove_cpu_columns(df)

def list_total_count(sp_list):
    '''
    Count the number of total items in a SharePoint list.
    '''
    all_items = sp_list.items.get_all(5000, print_progress).execute_query()
    print('Total items count: {0}'.format(len(all_items)))