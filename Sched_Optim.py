# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import datetime
import time

#import seaborn as sns

DATA_MASTER = pd.read_excel('C:/Users/arose/OneDrive - Ortho Molecular Products/Desktop/Schedule_Optimization/CurrentFormulations.xlsx', sheet_name='Export')

def create_empty_room_schedule():
    base_time = datetime.datetime.strptime("04:00:00", "%H:%M:%S")
    time_list = [(base_time + datetime.timedelta(minutes=x)).strftime("%H:%M") for x in range(20 * 60)]
    
    # Create a DataFrame with the time list and other columns initialized to 'Idle'
    room_schedule = pd.DataFrame({
        'time': time_list,
        'North': 'Idle',
        'South': 'Idle',
        'Old': 'Idle',
        'New': 'Idle'
    })
    
    return room_schedule

def prep_data(data_master):
    # pip install pandas numpy
    
    def get_mix_room(batch_size):
        # Generate random choices for each condition
        random_choices = np.random.choice(['North', 'South', 'Old', 'New'], size=len(batch_size), p=[0.25, 0.25, 0.25, 0.25])
        mix_room = np.where(batch_size.str.contains('1000', regex=False), 
                            np.where(random_choices[:len(batch_size)] == 'North', 'North', 'South'),
                            np.where(batch_size.str.contains('100', regex=False), 
                                     np.where(random_choices[:len(batch_size)] == 'North', 'North', 'South'),
                                     np.where(batch_size.str.contains('5000', regex=False), 
                                              np.where(random_choices[:len(batch_size)] == 'Old', 'Old', 'New'),
                                              np.where(batch_size.str.contains('500', regex=False), 
                                                       np.where(random_choices[:len(batch_size)] == 'North', 'North', 'South'),
                                                       np.where(batch_size.str.contains('200', regex=False), 
                                                                np.where(random_choices[:len(batch_size)] == 'North', 'North', 'South'),
                                                                np.nan)))))
        return mix_room
    
    def get_mill_room(mix_room):
        random_choices = np.random.choice(['North', 'South'], size=len(mix_room))
        mill_room = np.where((mix_room == 'North') | (mix_room == 'South'), mix_room,
                             np.where((mix_room == 'Old') | (mix_room == 'New'), 
                                      np.where(random_choices[:len(mix_room)] == 'North', 'North', 'South'), 
                                      np.nan))
        return mill_room
    
    # Apply the decision logic to the data_master data frame
    data_master['mixRoom'] = get_mix_room(data_master['BatchSize'])
    data_master['millRoom'] = get_mill_room(data_master['mixRoom'])
    
    # Generate run order and arrange the data frame by this order
    data_master['runOrder'] = np.random.permutation(len(data_master))
    return data_master.sort_values(by='runOrder')


def sub5000_sched(df, data, i):
    # Assuming 'data' is a pandas DataFrame that has been set up similarly to setDT(data) in R,
    # and 'i' is the index of the current row being processed in 'data'.

    # =============================================================================
    # i = 2
    # df = data.iloc[[i]]
    # data = DATA_MASTER
    # =============================================================================    
    clean_group = df[0,21]
    item_no = df[0,5]
    prod_no = df[0,4]
    setup_time = df[0,9]
    #milling_time = int(np.floor(0.3 * df.iloc[0]['Mixing Time']))
    milling_time = int(np.floor(0.3 * df[0,10]))
    #mixing_time = int(np.ceil(0.7 * df.iloc[0]['Mixing Time']))
    mixing_time = int(np.ceil(0.7 * df[0,10]))
    dry_clean_setup_time = df[0,17]
    wet_clean_setup_time = df[0,18]
    dry_clean_time = df[0,13]
    wet_clean_time = df[0,14]
    dry_clean_reset = df[0,15]
    wet_clean_reset = df[0,16] 

    if (clean_group == 'A') & (next_item_group == clean_group):
        total_time = (["{} - Setup - {} - {}".format(clean_group, item_no, prod_no)] * setup_time +
                      ["{} - Milling - {} - {}".format(clean_group, item_no, prod_no)] * milling_time +
                      ["{} - Mixing - {} - {}".format(clean_group, item_no, prod_no)] * mixing_time +
                      ["{} - Dry Clean Setup - {} - {}".format(clean_group, item_no, prod_no)] * dry_clean_setup_time +
                      ["{} - Dry Clean - {} - {}".format(clean_group, item_no, prod_no)] * dry_clean_time +
                      ["{} - Dry Clean teardown - {} - {}".format(clean_group, item_no, prod_no)] * dry_clean_reset)

        return total_time
    if (clean_group != 'A') | (next_item_group != clean_group):
        total_time = (["{} - Setup - {} - {}".format(clean_group, item_no, prod_no)] * setup_time + 
                      ["{} - Milling - {} - {}".format(clean_group, item_no, prod_no)] * milling_time +
                      ["{} - Mixing - {} - {}".format(clean_group, item_no, prod_no)] * mixing_time +
                      ["{} - Wet Clean Setup - {} - {}".format(clean_group, item_no, prod_no)] * wet_clean_setup_time +
                      ["{} - Wet Clean - {} - {}".format(clean_group, item_no, prod_no)] * wet_clean_time +
                      # Assuming 90 represents drying time
                      ["{} - Drying - {} - {}".format(clean_group, item_no, prod_no)] * 90 +
                      ["{} - Wet Clean teardown - {} - {}".format(clean_group, item_no, prod_no)] * wet_clean_reset)
        return total_time
#test = sub5000_sched(data.head(2), DATA_MASTER, 1)

def sup5000_sched(df, data, i):
    
    # =============================================================================
    #i = 2
    #df = data.iloc[[i]]
    #data = DATA_MASTER
    # =============================================================================

    
    mixing_room = df[0,22]      
    milling_room = df[0,23]
    clean_group = df[0,21]
    # Get the next item's group for the same mixing room
    data_array = data.to_numpy()

    # Define the column indices based on your note
    group_col_index = 21
    mix_room_col_index = 22
    mill_room_col_index = 23
    run_order_col_index = 24
    
    # Define default value for group if no next item is found
    default_group = 'C'
    
    # Find the next item for mixing room
    next_mix_mask = (data_array[:, mix_room_col_index] == mixing_room) & (data_array[:, run_order_col_index] > i)
    next_mix_indices = np.where(next_mix_mask)[0]
    next_item_group = data_array[next_mix_indices[0], group_col_index] if next_mix_indices.size > 0 else default_group
    
    # Find the next item for milling room
    next_mill_mask = (data_array[:, mill_room_col_index] == milling_room) & (data_array[:, run_order_col_index] > i)
    next_mill_indices = np.where(next_mill_mask)[0]
    next_mill_group = data_array[next_mill_indices[0], group_col_index] if next_mill_indices.size > 0 else default_group
        
# =============================================================================
#     next_item_df = data[(data['mixRoom'] == mixing_room) & (data['runOrder'] > i)].head(1)
#     next_item_group = next_item_df['Group'].iloc[0] if not next_item_df.empty else 'C'        
#     # Get the next item's group for the same milling room
#     next_mill_df = data[(data['millRoom'] == milling_room) & (data['runOrder'] > i)].head(1)
#     next_mill_group = next_mill_df['Group'].iloc[0] if not next_mill_df.empty else 'C'
# =============================================================================
    
    item_no = df[0,5]
    prod_no = df[0,4]
    setup_time = df[0,9]
    #milling_time = int(np.floor(0.3 * df.iloc[0]['Mixing Time']))
    milling_time = int(np.floor(0.3 * df[0,10]))
    #mixing_time = int(np.ceil(0.7 * df.iloc[0]['Mixing Time']))
    mixing_time = int(np.ceil(0.7 * df[0,10]))
    dry_clean_setup_time = df[0,17]
    wet_clean_setup_time = df[0,18]
    dry_clean_time = df[0,13]
    wet_clean_time = df[0,14]
    dry_clean_reset = df[0,15]
    wet_clean_reset = df[0,16]   
    mixing_room_total_time = []
    milling_room_total_time = []

    mixing_room_total_time = []
    milling_room_total_time = []

    def create_schedule(clean_group, prefix, item_no, prod_no, duration):
        return [f"{clean_group} - {prefix} - {item_no} - {prod_no}"] * duration

    if (clean_group == 'A') & (next_item_group == 'A'):
        mixing_room_total_time.extend(create_schedule(clean_group,"Setup", item_no, prod_no, setup_time))
        mixing_room_total_time.extend(create_schedule(clean_group,"Waiting - Milling", item_no, prod_no, milling_time))
        mixing_room_total_time.extend(create_schedule(clean_group,"Mixing", item_no, prod_no, mixing_time))
        mixing_room_total_time.extend(create_schedule(clean_group,"Dry Clean Setup", item_no, prod_no, dry_clean_setup_time))
        mixing_room_total_time.extend(create_schedule(clean_group,"Dry Clean", item_no, prod_no, dry_clean_time))
        mixing_room_total_time.extend(create_schedule(clean_group,"Dry Clean teardown", item_no, prod_no, dry_clean_reset))

    if (clean_group == 'A') & (next_mill_group == 'A'):
        milling_room_total_time.extend(create_schedule(clean_group,"Setup", item_no, prod_no, setup_time))
        milling_room_total_time.extend(create_schedule(clean_group,"Milling", item_no, prod_no, milling_time))
        milling_room_total_time.extend(create_schedule(clean_group,"Dry Clean Setup", item_no, prod_no, dry_clean_setup_time))
        milling_room_total_time.extend(create_schedule(clean_group,"Dry Clean", item_no, prod_no, dry_clean_time))
        milling_room_total_time.extend(create_schedule(clean_group,"Dry Clean teardown", item_no, prod_no, dry_clean_reset))

    if (clean_group == 'C') | (next_item_group != 'A'):
        mixing_room_total_time.extend(create_schedule(clean_group,"Setup", item_no, prod_no, setup_time))
        mixing_room_total_time.extend(create_schedule(clean_group,"Waiting - Milling", item_no, prod_no, milling_time))
        mixing_room_total_time.extend(create_schedule(clean_group,"Mixing", item_no, prod_no, mixing_time))
        mixing_room_total_time.extend(create_schedule(clean_group,"Wet Clean Setup", item_no, prod_no, wet_clean_setup_time))
        mixing_room_total_time.extend(create_schedule(clean_group,"Wet Clean", item_no, prod_no, wet_clean_time))
        mixing_room_total_time.extend(create_schedule(clean_group,"Drying", item_no, prod_no, 90))  # Assuming 90 represents drying time
        mixing_room_total_time.extend(create_schedule(clean_group,"Wet Clean teardown", item_no, prod_no, wet_clean_reset))

    if (clean_group == 'C') | (next_mill_group != 'A'):
        milling_room_total_time.extend(create_schedule(clean_group,"Setup", item_no, prod_no, setup_time))
        milling_room_total_time.extend(create_schedule(clean_group,"Milling", item_no, prod_no, milling_time))
        milling_room_total_time.extend(create_schedule(clean_group,"Wet Clean Setup", item_no, prod_no, wet_clean_setup_time))
        milling_room_total_time.extend(create_schedule(clean_group,"Wet Clean", item_no, prod_no, wet_clean_time))
        milling_room_total_time.extend(create_schedule(clean_group,"Drying", item_no, prod_no, 90))  # Assuming 90 represents drying time
        milling_room_total_time.extend(create_schedule(clean_group,"Wet Clean teardown", item_no, prod_no, wet_clean_reset))

    return {
        'mixing_room_total_time': mixing_room_total_time,
        'milling_room_total_time': milling_room_total_time
    }
#test = sup5000_sched(data.head(1), DATA_MASTER, 1)

def createGraph(df):
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    from datetime import datetime, timedelta
    # Assuming best_room_schedule is defined and populated elsewhere
    df = best_room_schedule
    df['time'] = pd.to_datetime(df['time'], format='%H:%M')
    # Pivot the DataFrame
    df_long = df.melt(id_vars='time', var_name='Room', value_name='Operation')
    
    # Group by Operation and Room, and summarize to get the start and end time of each Operation
    df_summarized = (
        df_long.groupby(['Operation', 'Room'])
        .agg(start_time=('time', 'min'), end_time=('time', 'max'))
        .reset_index()
    )
    # Add 60 seconds to end_time
    df_summarized['end_time'] = df_summarized['end_time'] + timedelta(seconds=60)
    # Add mid_time and mix_room
    df_summarized['mid_time'] = df_summarized['start_time'] + (df_summarized['end_time'] - df_summarized['start_time']) / 2
    df_summarized['mix_room'] = df_summarized['Room'].map({'North': 1, 'South': 2, 'Old': 3, 'New': 4})
    # Sort values
    df_summarized = df_summarized.sort_values(['mix_room', 'start_time'], ascending=[True, True])
    # Reset the index after all transformations
    df_summarized.reset_index(drop=True, inplace=True)
    df_summarized['PO'] = df_summarized['Operation'].str[-6:]


    def color_mapping(df):
        df = df_summarized
        df['color'] = 'blue'
        import random
        def generate_hex_color():
            # Generate a random color by generating a random number between 0x0 and 0xFFFFFF
            # and then converting it to a hex string.
            return "#{:06x}".format(random.randint(0, 0xFFFFFF))
        
        # Generate and print a random hex color code
        generate_hex_color()
        df['PO'].unique()
        for i in df['PO'].unique():
            df.loc[df['PO'] == i, 'color'] = generate_hex_color()
        return(df)
      
    df_summarized = color_mapping(df_summarized)
            

        
    #df_summarized = df_summarized
    
    fig, ax = plt.subplots(figsize=(20,16))
    # Create a Gantt chart with vertical bars
    for i, task in df_summarized.iterrows():
        #print(task['end_time'] - task['start_time'] )
        ax.bar(x = task['mix_room'],  # x-axis - Room position
               height = (task['end_time'] - task['start_time']),  # Bar height - Duration
               bottom = task['start_time'],  # Start time for the bottom of the bar
               width=0.8,  # Bar width
               align='center',
               color = task['color'],
               #color='skyblue',
               edgecolor='black')
        ax.text(task['mix_room'],  # x-axis - Room position
                task['mid_time'],  # Mid time for text position
                task['Operation'],  # The text inside the bar
                va='center',  # Vertical alignment
                ha='center',  # Horizontal alignment
                color='black')
    start_time = mdates.date2num(datetime.strptime("00:00:00", "%H:%M:%S"))
    end_time = mdates.date2num(datetime.strptime("04:00:00", "%H:%M:%S")) 
    start_time = -25566.833333333332  # Corresponds to 1900-01-01 04:00:00
    end_time = -25566.0 
    
    ax.set_ylim(start_time, end_time)
    # Set the x-ticks to room names
    room_ticks = df_summarized['mix_room'].unique()
    ax.set_xticks(room_ticks)
    ax.set_xticklabels(df_summarized['Room'].unique())
    # Format the y-ticks as dates
    ax.yaxis_date()
    ax.yaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
    ax.yaxis.set_major_locator(mdates.MinuteLocator(byminute=[0, 30]))  # Only show labels every half hour
    #ax.xaxis.set_major_locator(mdates.HourLocator(byhour=[0,1]))
    # Labeling Axes
    ax.set_xlabel('Room')
    ax.set_ylabel('Time')
    ax.set_title('Gantt Chart')
    plt.tight_layout()
    plt.show()



data = prep_data(DATA_MASTER)
best_batch_count = 0
best_active_minutes = 0
best_room_schedule = create_empty_room_schedule()
must_run = ['394742', '397001', '399332']
run_time = []

nrow = len(data)
nruns = 10000
i_run_time = 0
variable_time = 0 
next_room_time = 0
create_vector_time = 0
sub_5k_time = 0
sup_5k_time = 0
outer_loop_time = 0



for ii in range(nruns):
    start_time = time.time()
    batch_count = 0
    data = prep_data(DATA_MASTER)
    room_schedule = create_empty_room_schedule()
    PO_completed = []
    outer_loop_time = outer_loop_time + (time.time() - start_time)
    active_minutes = 0
    
    i_run_time_start = time.time()
    for i in range(nrow):
        start_time = time.time()
        df = data.iloc[[i]]
        df = df.to_numpy()

        
        #Vector Speed trial
        #################################################################################
        mixing_room = df[0,22]      
        milling_room = df[0,23]
        clean_group = df[0,21]
        
        next_room_time_start = time.time()        
        data_array = data.to_numpy()

        # Define the column indices based on your note
        group_col_index = 21
        mix_room_col_index = 22
        mill_room_col_index = 23
        run_order_col_index = 24
        
        # Define default value for group if no next item is found
        default_group = 'C'
        
        # Find the next item for mixing room
        next_mix_mask = (data_array[:, mix_room_col_index] == mixing_room) & (data_array[:, run_order_col_index] > i)
        next_mix_indices = np.where(next_mix_mask)[0]
        next_item_group = data_array[next_mix_indices[0], group_col_index] if next_mix_indices.size > 0 else default_group
        
        # Find the next item for milling room
        next_mill_mask = (data_array[:, mill_room_col_index] == milling_room) & (data_array[:, run_order_col_index] > i)
        next_mill_indices = np.where(next_mill_mask)[0]
        next_mill_group = data_array[next_mill_indices[0], group_col_index] if next_mill_indices.size > 0 else default_group
       
        next_room_time = next_room_time + (time.time() - next_room_time_start)       
        
        item_no = df[0,5]
        prod_no = df[0,4]
        setup_time = df[0,9]
        milling_time = int(np.floor(0.3 * df[0,10]))
        mixing_time = int(np.ceil(0.7 * df[0,10]))
        dry_clean_setup_time = df[0,17]
        wet_clean_setup_time = df[0,18]
        dry_clean_time = df[0,13]
        wet_clean_time = df[0,14]
        dry_clean_reset = df[0,15]
        wet_clean_reset = df[0,16]   
        mixing_room_total_time = []
        milling_room_total_time = []
        variable_time = variable_time + (time.time() - start_time)
       
        #logic to create a the schedule as a vector (2 if 5000l)
        #################################################################################
        start_time = time.time()
        if df[0,11] != 5000:
            total_time = sub5000_sched(df, data, i)
            
        if df[0,11] == 5000:
            total_time = sup5000_sched(df, data, i)            
            milling_room_total_time = total_time.get("milling_room_total_time")
            mixing_room_total_time = total_time.get("mixing_room_total_time")

        create_vector_time = create_vector_time + (time.time() - start_time)

            
        #logic to add sub <5000l POs to the room_schedule
        #################################################################################
        start_time = time.time()
        if df[0,11] != 5000:
            first_idle_time = room_schedule[mixing_room].eq('Idle').idxmax()
            if (len(room_schedule) - first_idle_time + 1) > len(total_time) and room_schedule[mixing_room].eq('Idle').any():
                run_length = len(total_time)
                room_schedule.loc[first_idle_time:(run_length + first_idle_time - 1), mixing_room] = total_time
                batch_count += 1
                PO_completed.append(prod_no)
                active_minutes = active_minutes + run_length
                #print('sub 5000', active_minutes)
# =============================================================================
#                 print('1000L Mixing: ' , list(total_time)[0][-6:],'-', mixing_room, 
#                       '\nStart Minute:  ', first_idle_time,
#                       '\nEnd Minute:  ', (run_length + first_idle_time - 1),
#                       '\n-----------------------------------------------')
# =============================================================================
        sub_5k_time = sub_5k_time + (time.time() - start_time)
        
        #Logic to add >5000l POs to the room_schedule dataframe
        #################################################################################
        first_idle_time_mixing_room = (room_schedule[mixing_room] == 'Idle').idxmax()
        first_idle_time_milling_room = (room_schedule[milling_room] == 'Idle').idxmax()
        
        # Ensure that a product wont be mixed before it is milled
        if first_idle_time_milling_room > first_idle_time_mixing_room:
            first_idle_time_mixing_room = first_idle_time_milling_room
        
        start_time = time.time()
        # Check the condition before updating the schedule.
        if (df[0,11] == 5000 and
                room_schedule[milling_room].eq('Idle').any() and
                room_schedule[mixing_room].eq('Idle').any() and
                (len(room_schedule) - first_idle_time_milling_room) > len(milling_room_total_time) and
                (len(room_schedule) - first_idle_time_mixing_room) > len(mixing_room_total_time)):
        
            run_length = len(milling_room_total_time)
            active_minutes = active_minutes + run_length
            #print('mill 5k', active_minutes)
            room_schedule.loc[first_idle_time_milling_room:first_idle_time_milling_room + run_length - 1, milling_room] = milling_room_total_time
# =============================================================================
#             print('5000L Milling: ' , milling_room_total_time[0][-6:],'-', milling_room, 
#                   '\nStart Minute:  ', first_idle_time_milling_room,
#                   '\nEnd Minute:  ', first_idle_time_milling_room + run_length - 1)
# =============================================================================
            
            run_length = len(mixing_room_total_time)
            active_minutes = active_minutes + run_length
            #print('mix 5k', active_minutes)

            room_schedule.loc[first_idle_time_mixing_room:first_idle_time_mixing_room + run_length - 1, mixing_room] = mixing_room_total_time
# =============================================================================
#             print('5000L Mixing: ' , mixing_room_total_time[0][-6:],'-', mixing_room, 
#                   '\nStart Minute:  ', first_idle_time_mixing_room,
#                   '\nEnd Minute:  ', first_idle_time_mixing_room + run_length - 1,
#                   '\n-----------------------------------------------')
# =============================================================================
            
            # Update the 'Idle' cells with "Idle i" in the mixing and milling schedules.
            
            # Create a boolean mask with the same index as room_schedule
            mask = (room_schedule[mixing_room].iloc[0:first_idle_time_mixing_room] == 'Idle').reindex(room_schedule.index, fill_value=False)            
            # Use the boolean mask to perform the assignment
            room_schedule.loc[mask, mixing_room] = f"Idle {i}"
            
            mask = (room_schedule[milling_room].iloc[0:first_idle_time_milling_room] == 'Idle').reindex(room_schedule.index, fill_value=False)            
            # Use the boolean mask to perform the assignment
            room_schedule.loc[mask, milling_room] = f"Idle {i}"
 
            # Update batch count and PO_completed list
            batch_count += 1
            PO_completed.append(prod_no)
        
        #print('Iter done############')
        sup_5k_time = sup_5k_time + (time.time() - start_time)

    i_run_time = i_run_time + (time.time() - i_run_time_start)
    
    
    start_time = time.time()
    
    #idle_count = (room_schedule.applymap(lambda x: "Idle" in str(x))).sum().sum()
    #idle_count = np.sum(np.char.find(room_schedule.values.astype(str), "Idle") != -1)


    outer_loop_time = outer_loop_time + (time.time() - start_time)
    
    if ((batch_count > best_batch_count) or
        (batch_count == best_batch_count and active_minutes < best_active_minutes) and
        all(np.isin(np.array(must_run).astype(int), np.array(PO_completed).astype(int)))):

        print(f"""
        ----------------------------------------
        New Best Schedule Found at Iteration: {ii}
        ----------------------------------------
        - Batches Ran: {batch_count}
        - Previous Best Batches Ran: {best_batch_count}
        - Active Time (minutes): {active_minutes}
        - Previous Best Active Time (minutes): {best_active_minutes}
        - Purchase Orders Completed: 
        {PO_completed}
        ----------------------------------------
        """)
        best_batch_count = batch_count
        best_active_minutes = active_minutes
        best_room_schedule = room_schedule.copy()

    
print(i_run_time)
print(variable_time)
print(create_vector_time)
print(sub_5k_time)
print(sup_5k_time)
print(outer_loop_time)

createGraph(best_room_schedule)
#end

