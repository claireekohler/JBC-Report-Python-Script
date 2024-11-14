```python
import pandas as pd
import numpy as np
```


```python
df = pd.read_excel(r'/Users/claireekohler/Downloads/JBC_Test.xlsx')
```


```python
print(df)
```

             Job Phase  Segment    Job-Ph-Seg  S/S Asgn Firm Asgn Div Task ID  \
    0    73461.0    BL      1.0  73461-BL-001  1.0        HI       NW    E200   
    1    73461.0    BL      1.0  73461-BL-001  1.0        HI       NW    P300   
    2    73461.0    BL      1.0  73461-BL-001  1.0        HI       NW    P902   
    3    73461.0    BL      1.0  73461-BL-001  1.0        HI       HG    E200   
    4    73461.0    BL      1.0  73461-BL-001  1.0        HI       HG    P300   
    ..       ...   ...      ...           ...  ...       ...      ...     ...   
    347  73461.0    PL     31.0  73461-PL-031  1.0        HI       NW    E999   
    348  73461.0    PL     31.0  73461-PL-031  1.0        HI       NW    T901   
    349  73461.0    PL     31.0  73461-PL-031  2.0        HI       NW    E005   
    350  73461.0    PL     31.0  73461-PL-031  3.0        HI       NW    C999   
    351      NaN   NaN      NaN           NaN  NaN       NaN      NaN     NaN   
    
         Task Asgn Ofc S/S S/L  ...                   Segment Title  \
    0              4.0      RL  ...  Marketing                        
    1              4.0      MA  ...  Marketing                        
    2              4.0      MA  ...  Marketing                        
    3             12.0      RL  ...  Marketing                        
    4             12.0      MA  ...  Marketing                        
    ..             ...     ...  ...                             ...   
    347            4.0      RL  ...  TO 31 : Process Update & Repor   
    348            4.0      TR  ...  TO 31 : Process Update & Repor   
    349            4.0      RL  ...  TO 31 : Process Update & Repor   
    350            4.0      CS  ...  TO 31 : Process Update & Repor   
    351            NaN     NaN  ...                             NaN   
    
                       Subsegment Title    Sub # Bud Fee Type  Primary Client  \
    0    Marketing                           NaN          0.0   SOUND TRANSIT   
    1    Marketing                           NaN          0.0   SOUND TRANSIT   
    2    Marketing                           NaN          0.0   SOUND TRANSIT   
    3    Marketing                           NaN          0.0   SOUND TRANSIT   
    4    Marketing                           NaN          0.0   SOUND TRANSIT   
    ..                              ...      ...          ...             ...   
    347  Task 1-Project Management           NaN          1.0   SOUND TRANSIT   
    348  Task 1-Project Management           NaN          1.0   SOUND TRANSIT   
    349  Subconsultants                  55868.0          1.0   SOUND TRANSIT   
    350  PM Contingency                      NaN          1.0   SOUND TRANSIT   
    351                             NaN      NaN          NaN             NaN   
    
        Primary Client Loc Secondary Client Secondary Client Loc  \
    0                  1.0               --                  ---   
    1                  1.0               --                  ---   
    2                  1.0               --                  ---   
    3                  1.0               --                  ---   
    4                  1.0               --                  ---   
    ..                 ...              ...                  ...   
    347                1.0               --                  ---   
    348                1.0               --                  ---   
    349                1.0               --                  ---   
    350                1.0               --                  ---   
    351                NaN              NaN                  NaN   
    
             Market Sector  HNTB Advantage  
    0    RAIL/TRANSIT-Muni               N  
    1    RAIL/TRANSIT-Muni               N  
    2    RAIL/TRANSIT-Muni               N  
    3    RAIL/TRANSIT-Muni               N  
    4    RAIL/TRANSIT-Muni               N  
    ..                 ...             ...  
    347  RAIL/TRANSIT-Muni               N  
    348  RAIL/TRANSIT-Muni               N  
    349  RAIL/TRANSIT-Muni               N  
    350  RAIL/TRANSIT-Muni               N  
    351                NaN             NaN  
    
    [352 rows x 51 columns]



```python
hours_index = df.columns.get_loc('Hours')
print(hours_index)
```

    19



```python
#Add 'Adjustment' and 'New Hours' Columns

df.insert(hours_index + 1, 'Adjustment', 0)
df.insert(hours_index + 2, 'New Hours',0)
```


```python
#Check that the columns have been added
print(df[['Hours','Adjustment','New Hours']])
```

         Hours  Adjustment  New Hours
    0        0         NaN        NaN
    1       45         NaN        NaN
    2       10         NaN        NaN
    3        0         NaN        NaN
    4        5         NaN        NaN
    ..     ...         ...        ...
    347      0         NaN        NaN
    348     88         NaN        NaN
    349      0         NaN        NaN
    350     84         NaN        NaN
    351   6074         NaN        NaN
    
    [352 rows x 3 columns]



```python
# Convert 'Job-Ph-Seg' and 'S/S' columns to string type
df['Job-Ph-Seg'] = df['Job-Ph-Seg'].astype(str)
df['S/S'] = df['S/S'].astype(str)
```


```python
df.dtypes
```




    Job                              float64
    Phase                             object
    Segment                          float64
    Job-Ph-Seg                        object
    S/S                              float64
    Asgn Firm                         object
    Asgn Div                          object
    Task ID                           object
    Task Asgn Ofc                    float64
    S/S S/L                           object
    Task Assign Name                  object
    PM Name                           object
    Owning Firm                       object
    Owning Div                        object
    PM Ofc                           float64
    Task Status                       object
    Task Type                         object
    Task Description                  object
    Baseline ER                      float64
    Hours                              int64
    Adjustment                       float64
    New Hours                        float64
    Rate                             float64
    Base Cost                          int64
    Bud ER                           float64
    Loaded Budgeted Cost               int64
    Fixed Fee Cap                    float64
    MTD Hours                          int64
    MTD Costs                          int64
    YTD Hours                        float64
    YTD Costs                        float64
    JTD Hours                        float64
    JTD Rate                         float64
    JTD Costs                        float64
    Loaded JTD Costs                 float64
    Loaded Costs Remaining           float64
    % Hours                          float64
    % Costs                          float64
    Actual Remain Hours              float64
    Actual Remain Costs              float64
    Costs Last Posted         datetime64[ns]
    Last Updated Date         datetime64[ns]
    Job Title                         object
    Segment Title                     object
    Subsegment Title                  object
    Sub #                            float64
    Bud Fee Type                     float64
    Primary Client                    object
    Primary Client Loc               float64
    Secondary Client                  object
    Secondary Client Loc              object
    Market Sector                     object
    HNTB Advantage                    object
    dtype: object




```python
#Function to add hours to 'Adjustment' and 'New Hours'

#s/s cant be 999 or 997

def cover_hours(df):
    # Group the DataFrame by 'Job-Ph-Seg'
    grouped = df.groupby('Job-Ph-Seg')

    # Loop through each group
    for name, group in grouped:
        for idx, row in group.iterrows():
            # Check if 'Job-Ph-Seg' contains 'PL' and 'S/S' does not contain '999'
            if 'PL' in row['Job-Ph-Seg'] and '999' not in row['S/S']:
                # Get the value of 'Actual Remain Hours'
                actual_remaining_hours = row['Actual Remain Hours']
                
                # If 'Actual Remain Hours' is negative
                if actual_remaining_hours < 0:
                    # Calculate the adjustment as |Actual Remaining Hours| + 1
                    X = abs(actual_remaining_hours) + 1
                    
                    # Set 'Adjustment' to X and update 'New Hours'
                    df.at[idx, 'Adjustment'] = X
                    df.at[idx, 'New Hours'] = row['Hours'] + X     
    return df
```


```python
#Function to balance hours

def balance_hours(df):
    #Group by 'Job-Ph-Seg' to handle each Ph-Seg independently
    grouped = df.groupby('Job-Ph-Seg')

    #Loop through each group
    for name, group in grouped:
        #Loop through each row in current segment 
        for idx, row in group.iterrows():
            #Consider rows only where adjusted
            adjustment = row['Adjustment']
            if adjustment > 0:
                #Find rows in same segment with 'Actual Remain Hours' greater than 'Adjustment'
                eligible_rows = group[(group['Actual Remain Hours'] > adjustment) & (group['Adjustment'] == 0)]
                if not eligible_rows.empty:
                   target_idx = eligible_rows.index[0]
                   df.at[target_idx, 'Adjustment'] = -adjustment
                   df.at[target_idx,'New Hours'] = df.at[target_idx, 'Hours'] + df.at[target_idx, 'Adjustment']

    return df

```


```python
#Add 'new loaded costs remaining' and 'net' columns

loaded_costs_remain_index = df.columns.get_loc('Loaded Costs Remaining')

df.insert(loaded_costs_remain_index + 1, 'New Loaded Costs Remaining', 0)
df.insert(loaded_costs_remain_index + 2, 'NET',0)
```


```python
#Function to calculate 'New Loaded Costs' Remaining and 'NET'

#if 'new hours' isnt null,
#insert into 'new loaded costs remaining' = (new hours*rate*budget ER)-loaded costs remaining
#net = adjustment*rate*budgeted ER

def new_loaded_costs(df):
    # Group by 'Job-Ph-Seg' to handle each Ph-Seg independently
    grouped = df.groupby('Job-Ph-Seg')

    # Loop through each group
    for name, group in grouped:
        # Loop through each row in the current segment
        for idx, row in group.iterrows():
            # Consider only rows with a non-null 'Adjustment'
            adjustment = row['Adjustment']
            if adjustment != 0:
                # Calculate 'New Loaded Costs Remaining'
                df.at[idx, 'New Loaded Costs Remaining'] = (
                    df.at[idx, 'New Hours'] * df.at[idx, 'Bud ER'] * df.at[idx, 'Rate']
                    - df.at[idx, 'Loaded Costs Remaining']
                )
                # Calculate 'NET'
                df.at[idx, 'NET'] = (
                    df.at[idx, 'Adjustment'] * df.at[idx, 'Bud ER'] * df.at[idx, 'Rate']
                )

    return df


```


```python
#Apply functions to dataframe and view output

df = cover_hours(df)
```

             Job Phase  Segment    Job-Ph-Seg  S/S Asgn Firm Asgn Div Task ID  \
    0    73461.0    BL      1.0  73461-BL-001  1.0        HI       NW    E200   
    1    73461.0    BL      1.0  73461-BL-001  1.0        HI       NW    P300   
    2    73461.0    BL      1.0  73461-BL-001  1.0        HI       NW    P902   
    3    73461.0    BL      1.0  73461-BL-001  1.0        HI       HG    E200   
    4    73461.0    BL      1.0  73461-BL-001  1.0        HI       HG    P300   
    ..       ...   ...      ...           ...  ...       ...      ...     ...   
    347  73461.0    PL     31.0  73461-PL-031  1.0        HI       NW    E999   
    348  73461.0    PL     31.0  73461-PL-031  1.0        HI       NW    T901   
    349  73461.0    PL     31.0  73461-PL-031  2.0        HI       NW    E005   
    350  73461.0    PL     31.0  73461-PL-031  3.0        HI       NW    C999   
    351      NaN   NaN      NaN           nan  nan       NaN      NaN     NaN   
    
         Task Asgn Ofc S/S S/L  ...                   Segment Title  \
    0              4.0      RL  ...  Marketing                        
    1              4.0      MA  ...  Marketing                        
    2              4.0      MA  ...  Marketing                        
    3             12.0      RL  ...  Marketing                        
    4             12.0      MA  ...  Marketing                        
    ..             ...     ...  ...                             ...   
    347            4.0      RL  ...  TO 31 : Process Update & Repor   
    348            4.0      TR  ...  TO 31 : Process Update & Repor   
    349            4.0      RL  ...  TO 31 : Process Update & Repor   
    350            4.0      CS  ...  TO 31 : Process Update & Repor   
    351            NaN     NaN  ...                             NaN   
    
                       Subsegment Title    Sub # Bud Fee Type  Primary Client  \
    0    Marketing                           NaN          0.0   SOUND TRANSIT   
    1    Marketing                           NaN          0.0   SOUND TRANSIT   
    2    Marketing                           NaN          0.0   SOUND TRANSIT   
    3    Marketing                           NaN          0.0   SOUND TRANSIT   
    4    Marketing                           NaN          0.0   SOUND TRANSIT   
    ..                              ...      ...          ...             ...   
    347  Task 1-Project Management           NaN          1.0   SOUND TRANSIT   
    348  Task 1-Project Management           NaN          1.0   SOUND TRANSIT   
    349  Subconsultants                  55868.0          1.0   SOUND TRANSIT   
    350  PM Contingency                      NaN          1.0   SOUND TRANSIT   
    351                             NaN      NaN          NaN             NaN   
    
        Primary Client Loc Secondary Client Secondary Client Loc  \
    0                  1.0               --                  ---   
    1                  1.0               --                  ---   
    2                  1.0               --                  ---   
    3                  1.0               --                  ---   
    4                  1.0               --                  ---   
    ..                 ...              ...                  ...   
    347                1.0               --                  ---   
    348                1.0               --                  ---   
    349                1.0               --                  ---   
    350                1.0               --                  ---   
    351                NaN              NaN                  NaN   
    
             Market Sector  HNTB Advantage  
    0    RAIL/TRANSIT-Muni               N  
    1    RAIL/TRANSIT-Muni               N  
    2    RAIL/TRANSIT-Muni               N  
    3    RAIL/TRANSIT-Muni               N  
    4    RAIL/TRANSIT-Muni               N  
    ..                 ...             ...  
    347  RAIL/TRANSIT-Muni               N  
    348  RAIL/TRANSIT-Muni               N  
    349  RAIL/TRANSIT-Muni               N  
    350  RAIL/TRANSIT-Muni               N  
    351                NaN             NaN  
    
    [352 rows x 53 columns]



```python
df = balance_hours(df)
df = new_loaded_costs(df)
```


```python
print(df)
```

             Job Phase  Segment    Job-Ph-Seg  S/S Asgn Firm Asgn Div Task ID  \
    0    73461.0    BL      1.0  73461-BL-001  1.0        HI       NW    E200   
    1    73461.0    BL      1.0  73461-BL-001  1.0        HI       NW    P300   
    2    73461.0    BL      1.0  73461-BL-001  1.0        HI       NW    P902   
    3    73461.0    BL      1.0  73461-BL-001  1.0        HI       HG    E200   
    4    73461.0    BL      1.0  73461-BL-001  1.0        HI       HG    P300   
    ..       ...   ...      ...           ...  ...       ...      ...     ...   
    347  73461.0    PL     31.0  73461-PL-031  1.0        HI       NW    E999   
    348  73461.0    PL     31.0  73461-PL-031  1.0        HI       NW    T901   
    349  73461.0    PL     31.0  73461-PL-031  2.0        HI       NW    E005   
    350  73461.0    PL     31.0  73461-PL-031  3.0        HI       NW    C999   
    351      NaN   NaN      NaN           nan  nan       NaN      NaN     NaN   
    
         Task Asgn Ofc S/S S/L  ...    Sub # Bud Fee Type Primary Client  \
    0              4.0      RL  ...      NaN          0.0  SOUND TRANSIT   
    1              4.0      MA  ...      NaN          0.0  SOUND TRANSIT   
    2              4.0      MA  ...      NaN          0.0  SOUND TRANSIT   
    3             12.0      RL  ...      NaN          0.0  SOUND TRANSIT   
    4             12.0      MA  ...      NaN          0.0  SOUND TRANSIT   
    ..             ...     ...  ...      ...          ...            ...   
    347            4.0      RL  ...      NaN          1.0  SOUND TRANSIT   
    348            4.0      TR  ...      NaN          1.0  SOUND TRANSIT   
    349            4.0      RL  ...  55868.0          1.0  SOUND TRANSIT   
    350            4.0      CS  ...      NaN          1.0  SOUND TRANSIT   
    351            NaN     NaN  ...      NaN          NaN            NaN   
    
        Primary Client Loc  Secondary Client Secondary Client Loc  \
    0                  1.0                --                  ---   
    1                  1.0                --                  ---   
    2                  1.0                --                  ---   
    3                  1.0                --                  ---   
    4                  1.0                --                  ---   
    ..                 ...               ...                  ...   
    347                1.0                --                  ---   
    348                1.0                --                  ---   
    349                1.0                --                  ---   
    350                1.0                --                  ---   
    351                NaN               NaN                  NaN   
    
             Market Sector HNTB Advantage  New Loaded Costs Remaining  NET  
    0    RAIL/TRANSIT-Muni              N                         0.0  0.0  
    1    RAIL/TRANSIT-Muni              N                         0.0  0.0  
    2    RAIL/TRANSIT-Muni              N                         0.0  0.0  
    3    RAIL/TRANSIT-Muni              N                         0.0  0.0  
    4    RAIL/TRANSIT-Muni              N                         0.0  0.0  
    ..                 ...            ...                         ...  ...  
    347  RAIL/TRANSIT-Muni              N                       -54.0  0.0  
    348  RAIL/TRANSIT-Muni              N                    -12631.8  0.0  
    349  RAIL/TRANSIT-Muni              N                   -743417.0  0.0  
    350  RAIL/TRANSIT-Muni              N                    -14565.6  0.0  
    351                NaN            NaN                         NaN  NaN  
    
    [352 rows x 55 columns]



```python
df.to_excel('JBC_Calc.xlsx', index=False)
```


```python

```
