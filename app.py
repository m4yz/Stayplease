with tab3:

    st.subheader("📋 Task Explorer")

    # =====================================================
    # TASK EXPLORER FILTERS
    # =====================================================

    col1, col2, col3 = st.columns(3)

    with col1:

        explorer_teams = sorted(
            filtered["Team"]
            .dropna()
            .unique()
            .tolist()
        )

        selected_explorer_teams = st.multiselect(
            "🏢 Team",
            explorer_teams,
            default=explorer_teams,
            key="explorer_team"
        )


    with col2:

        explorer_locations = sorted(
            filtered["Location"]
            .dropna()
            .unique()
            .tolist()
        )

        selected_locations = st.multiselect(
            "📍 Location",
            explorer_locations,
            default=explorer_locations,
            key="explorer_location"
        )


    with col3:

        explorer_assignees = sorted(
            filtered["Assignee"]
            .dropna()
            .unique()
            .tolist()
        )

        selected_assignees = st.multiselect(
            "👤 Assignee",
            explorer_assignees,
            default=explorer_assignees,
            key="explorer_assignee"
        )


    col4, col5, col6 = st.columns(3)

    with col4:

        explorer_statuses = sorted(
            filtered["Status"]
            .dropna()
            .unique()
            .tolist()
        )

        selected_explorer_statuses = st.multiselect(
            "📌 Status",
            explorer_statuses,
            default=explorer_statuses,
            key="explorer_status"
        )


    with col5:

        explorer_priorities = sorted(
            filtered["Priority"]
            .dropna()
            .unique()
            .tolist()
        )

        selected_explorer_priorities = st.multiselect(
            "🔥 Priority",
            explorer_priorities,
            default=explorer_priorities,
            key="explorer_priority"
        )


    with col6:

        search = st.text_input(
            "🔎 Search",
            placeholder="Task, requestor, assignee...",
            key="explorer_search"
        )


    # =====================================================
    # APPLY TASK EXPLORER FILTERS
    # =====================================================

    display = filtered.copy()


    # TEAM

    if selected_explorer_teams:

        display = display[
            display["Team"].isin(
                selected_explorer_teams
            )
        ]


    # LOCATION

    if selected_locations:

        display = display[
            display["Location"].isin(
                selected_locations
            )
        ]


    # ASSIGNEE

    if selected_assignees:

        display = display[
            display["Assignee"].isin(
                selected_assignees
            )
        ]


    # STATUS

    if selected_explorer_statuses:

        display = display[
            display["Status"].isin(
                selected_explorer_statuses
            )
        ]


    # PRIORITY

    if selected_explorer_priorities:

        display = display[
            display["Priority"].isin(
                selected_explorer_priorities
            )
        ]


    # =====================================================
    # SEARCH
    # =====================================================

    if search:

        search_columns = [

            "Task",
            "Location",
            "Requestor",
            "Assignee",
            "Message",
            "Req Department",
            "Department"

        ]


        mask = pd.Series(
            False,
            index=display.index
        )


        for col in search_columns:

            mask |= (

                display[col]
                .fillna("")
                .astype(str)
                .str.contains(
                    search,
                    case=False,
                    na=False
                )

            )


        display = display[mask]


    # =====================================================
    # DISPLAY COLUMNS
    # =====================================================

    display_columns = [

        "Team",
        "Location",
        "Task",
        "Quantity",
        "Req Department",
        "Requestor",
        "Department",
        "Assignee",
        "Status",
        "Priority",
        "To Do",
        "Doing",
        "Done",
        "Duration Hours",
        "Pause Reason",
        "Message"

    ]


    # =====================================================
    # RESULT COUNTER
    # =====================================================

    st.caption(
        f"Showing {len(display):,} task(s)"
    )


    # =====================================================
    # DATA TABLE
    # =====================================================

    st.dataframe(

        display[display_columns]
        .sort_values(
            "To Do",
            ascending=False
        ),

        use_container_width=True,
        hide_index=True,
        height=600

    )


    # =====================================================
    # DOWNLOAD FILTERED DATA
    # =====================================================

    csv = (

        display[display_columns]
        .to_csv(
            index=False
        )
        .encode("utf-8")

    )


    st.download_button(

        "⬇️ Download Filtered Data (CSV)",

        data=csv,

        file_name="stayplease_filtered_tasks.csv",

        mime="text/csv"

    )
