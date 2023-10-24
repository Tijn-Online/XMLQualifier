import pandas as pd
import pm4py

if __name__ == "__main__":
    dataframe = pd.read_csv('basware_extract.csv', sep=',')
    print(dataframe)
    dataframe = pm4py.format_dataframe(dataframe, case_id='ROOT_DOCUMENT_ID', activity_key='RESOURCE_ID', timestamp_key='TIME_STAMP')
    log = pm4py.convert_to_event_log(dataframe)



