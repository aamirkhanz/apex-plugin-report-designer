-- Report Designer region plug-in (TechRevX, www.techrevx.com)
-- render: writes the designer container and a JSON block with the settings and the template list, then calls
--         techrevx.reportDesigner.init (report-designer.js). The report data is loaded with AJAX (DATA).
-- ajax:   x01 = DATA    rows of the region source: { columns: [ { name, type } ], rows: [ [ ... ] ], truncated }
--         x01 = LIST    the templates of this region's group the user can see (own and shared ones)
--         x01 = LOAD    one template (x02 = id)
--         x01 = SAVE    inserts (x02 empty) or updates an own template: x03 = name, x04 = description,
--                       x05 = shared (Y/N), p_clob_01 = template JSON. Designer mode only.
--         x01 = DELETE  deletes an own template (x02 = id). Designer mode only.
-- Templates table: TEMPLATE_ID (generated), TEMPLATE_GROUP, NAME, DESCRIPTION, TEMPLATE (JSON, CLOB), OWNER,
-- IS_SHARED (Y/N), CREATED_ON, UPDATED_ON. Its name comes from the Templates Table attribute and must be an
-- existing table or view (dbms_assert).

function table_name( p_region in apex_plugin.t_region ) return varchar2 is
begin
    return sys.dbms_assert.sql_object_name( nvl( trim( p_region.attributes.get_varchar2( 'TEMPLATE_TABLE' ) ), 'TX_REPORT_TEMPLATES' ) );
end table_name;

function group_name( p_region in apex_plugin.t_region ) return varchar2 is
begin
    return nvl( substr( p_region.attributes.get_varchar2( 'TEMPLATE_GROUP', p_do_substitutions => true ), 1, 255 ), 'REPORTS' );
end group_name;

-- templates belong to this name: the Owner attribute (default &APP_USER.)
function owner_name( p_region in apex_plugin.t_region ) return varchar2 is
begin
    return nvl( substr( p_region.attributes.get_varchar2( 'OWNER', p_do_substitutions => true ), 1, 255 ), v( 'APP_USER' ) );
end owner_name;

function can_design( p_region in apex_plugin.t_region ) return boolean is
begin
    return nvl( p_region.attributes.get_varchar2( 'MODE' ), 'designer' ) = 'designer';
end can_design;

procedure write_list( p_region in apex_plugin.t_region ) is
    l_cursor  sys_refcursor;
    l_id      number;
    l_name    varchar2(255);
    l_desc    varchar2(4000);
    l_owner   varchar2(255);
    l_shared  varchar2(1);
    l_updated varchar2(30);
    l_me      varchar2(255) := owner_name( p_region );
begin
    open l_cursor for
        'select template_id, name, description, owner, is_shared, to_char( updated_on, ''YYYY-MM-DD"T"HH24:MI:SS'' )'
        || '  from ' || table_name( p_region )
        || ' where template_group = :grp and ( owner = :me or is_shared = ''Y'' )'
        || ' order by case when owner = :me2 then 0 else 1 end, upper( name )'
        using group_name( p_region ), l_me, l_me;
    apex_json.open_array( 'templates' );
    loop
        fetch l_cursor into l_id, l_name, l_desc, l_owner, l_shared, l_updated;
        exit when l_cursor%notfound;
        apex_json.open_object;
        apex_json.write( 'id', l_id );
        apex_json.write( 'name', l_name );
        apex_json.write( 'description', l_desc );
        apex_json.write( 'shared', l_shared = 'Y' );
        apex_json.write( 'mine', l_owner = l_me );
        apex_json.write( 'owner', case when l_owner <> l_me then l_owner end );
        apex_json.write( 'updated', l_updated );
        apex_json.close_object;
    end loop;
    close l_cursor;
    apex_json.close_array;
exception
    when others then
        if l_cursor%isopen then
            close l_cursor;
        end if;
        raise;
end write_list;

procedure write_template( p_region in apex_plugin.t_region, p_id in number ) is
    l_cursor sys_refcursor;
    l_name   varchar2(255);
    l_desc   varchar2(4000);
    l_json   clob;
    l_owner  varchar2(255);
    l_shared varchar2(1);
    l_me     varchar2(255) := owner_name( p_region );
begin
    open l_cursor for
        'select name, description, template, owner, is_shared from ' || table_name( p_region )
        || ' where template_id = :id and template_group = :grp and ( owner = :me or is_shared = ''Y'' )'
        using p_id, group_name( p_region ), l_me;
    fetch l_cursor into l_name, l_desc, l_json, l_owner, l_shared;
    if l_cursor%notfound then
        close l_cursor;
        raise_application_error( -20001, 'Report Designer: the template was not found or is not shared with you.' );
    end if;
    close l_cursor;
    apex_json.open_object;
    apex_json.write( 'id', p_id );
    apex_json.write( 'name', l_name );
    apex_json.write( 'description', l_desc );
    apex_json.write( 'shared', l_shared = 'Y' );
    apex_json.write( 'mine', l_owner = l_me );
    apex_json.write( 'owner', case when l_owner <> l_me then l_owner end );
    apex_json.write( 'template', l_json );
    apex_json.close_object;
end write_template;

-- the region source as columns and rows; dates are sent as ISO text (YYYY-MM-DDTHH24:MI:SS)
procedure write_data( p_region in apex_plugin.t_region ) is
    l_context apex_exec.t_context;
    l_max     pls_integer := nvl( p_region.attributes.get_number( 'MAX_ROWS' ), 5000 );
    l_types   apex_t_varchar2 := apex_t_varchar2();
    l_count   pls_integer := 0;
    l_type    pls_integer;
begin
    l_context := apex_exec.open_query_context( p_max_rows => l_max + 1 );
    apex_json.open_object;
    apex_json.open_array( 'columns' );
    for i in 1 .. apex_exec.get_column_count( l_context ) loop
        l_type := apex_exec.get_column( l_context, i ).data_type;
        l_types.extend;
        l_types( i ) := case
                            when l_type = apex_exec.c_data_type_number then 'number'
                            when l_type in ( apex_exec.c_data_type_date, apex_exec.c_data_type_timestamp,
                                             apex_exec.c_data_type_timestamp_tz, apex_exec.c_data_type_timestamp_ltz ) then 'date'
                            when l_type = apex_exec.c_data_type_clob then 'clob'
                            when l_type in ( apex_exec.c_data_type_blob, apex_exec.c_data_type_bfile ) then 'binary'
                            else 'text'
                        end;
        continue when l_types( i ) = 'binary';
        apex_json.open_object;
        apex_json.write( 'name', apex_exec.get_column( l_context, i ).name );
        apex_json.write( 'type', case when l_types( i ) = 'clob' then 'text' else l_types( i ) end );
        apex_json.close_object;
    end loop;
    apex_json.close_array;

    apex_json.open_array( 'rows' );
    while apex_exec.next_row( l_context ) loop
        l_count := l_count + 1;
        exit when l_count > l_max;
        apex_json.open_array;
        for i in 1 .. l_types.count loop
            case l_types( i )
                when 'number' then
                    apex_json.write( apex_exec.get_number( l_context, i ) );
                when 'date' then
                    apex_json.write( to_char( case
                                                  when apex_exec.get_column( l_context, i ).data_type = apex_exec.c_data_type_date
                                                      then apex_exec.get_date( l_context, i )
                                                  else cast( apex_exec.get_timestamp( l_context, i ) as date )
                                              end, 'YYYY-MM-DD"T"HH24:MI:SS' ) );
                when 'clob' then
                    apex_json.write( dbms_lob.substr( apex_exec.get_clob( l_context, i ), 8000, 1 ) );
                when 'binary' then
                    null;
                else
                    apex_json.write( apex_exec.get_varchar2( l_context, i ) );
            end case;
        end loop;
        apex_json.close_array;
    end loop;
    apex_json.close_array;
    apex_json.write( 'truncated', l_count > l_max );
    apex_json.close_object;
    apex_exec.close( l_context );
exception
    when others then
        apex_exec.close( l_context );
        raise;
end write_data;

procedure save_template( p_region in apex_plugin.t_region ) is
    l_id     number := to_number( apex_application.g_x02 default null on conversion error );
    l_name   varchar2(255) := substr( trim( apex_application.g_x03 ), 1, 255 );
    l_desc   varchar2(4000) := substr( apex_application.g_x04, 1, 4000 );
    l_shared varchar2(1) := case when apex_application.g_x05 = 'Y' and p_region.attributes.get_varchar2( 'ALLOW_SHARING' ) = 'Y' then 'Y' else 'N' end;
    l_json   clob := apex_application.g_clob_01;
    l_check  json_object_t;
begin
    if not can_design( p_region ) then
        raise_application_error( -20002, 'Report Designer: templates can only be saved in Designer mode.' );
    end if;
    if l_name is null then
        raise_application_error( -20003, 'Report Designer: enter a name for the template.' );
    end if;
    l_check := json_object_t.parse( l_json );   -- raises an error for anything that is not a JSON object

    if l_id is null then
        execute immediate
            'insert into ' || table_name( p_region )
            || ' ( template_group, name, description, template, owner, is_shared, created_on, updated_on )'
            || ' values ( :grp, :name, :descr, :tpl, :me, :shared, systimestamp, systimestamp )'
            || ' returning template_id into :id'
            using group_name( p_region ), l_name, l_desc, l_json, owner_name( p_region ), l_shared
            returning into l_id;
    else
        execute immediate
            'update ' || table_name( p_region )
            || '   set name = :name, description = :descr, template = :tpl, is_shared = :shared, updated_on = systimestamp'
            || ' where template_id = :id and template_group = :grp and owner = :me'
            using l_name, l_desc, l_json, l_shared, l_id, group_name( p_region ), owner_name( p_region );
        if sql%rowcount = 0 then
            raise_application_error( -20004, 'Report Designer: you can only change your own templates. Use Save As to keep a copy.' );
        end if;
    end if;

    apex_json.open_object;
    apex_json.write( 'id', l_id );
    write_list( p_region );
    apex_json.close_object;
end save_template;

procedure delete_template( p_region in apex_plugin.t_region ) is
begin
    if not can_design( p_region ) then
        raise_application_error( -20002, 'Report Designer: templates can only be deleted in Designer mode.' );
    end if;
    execute immediate
        'delete from ' || table_name( p_region ) || ' where template_id = :id and template_group = :grp and owner = :me'
        using to_number( apex_application.g_x02 default null on conversion error ), group_name( p_region ), owner_name( p_region );
    if sql%rowcount = 0 then
        raise_application_error( -20005, 'Report Designer: you can only delete your own templates.' );
    end if;
    apex_json.open_object;
    write_list( p_region );
    apex_json.close_object;
end delete_template;

procedure render (
    p_plugin in            apex_plugin.t_plugin,
    p_region in            apex_plugin.t_region,
    p_param  in            apex_plugin.t_region_render_param,
    p_result in out nocopy apex_plugin.t_region_render_result )
is
    l_json clob;
    l_item varchar2(255) := p_region.attributes.get_varchar2( 'TEMPLATE_ITEM' );
begin
    apex_json.initialize_clob_output;
    apex_json.open_object;
    apex_json.write( 'ajaxIdentifier', apex_plugin.get_ajax_identifier );
    apex_json.write( 'itemsToSubmit', apex_plugin_util.page_item_names_to_jquery( p_region.ajax_items_to_submit ) );
    apex_json.write( 'noDataFound', apex_plugin_util.replace_substitutions( p_region.no_data_found_message, false ) );
    apex_json.write( 'filePrefix', p_plugin.file_prefix );
    apex_json.write( 'mode', nvl( p_region.attributes.get_varchar2( 'MODE' ), 'designer' ) );
    apex_json.write( 'formats', p_region.attributes.get_varchar2( 'FORMATS' ) );
    apex_json.write( 'sharing', p_region.attributes.get_varchar2( 'ALLOW_SHARING' ) = 'Y' );
    apex_json.write( 'user', v( 'APP_USER' ) );
    apex_json.write( 'templateItem', l_item );
    if l_item is not null then
        apex_json.write( 'templateId', v( l_item ) );
    end if;
    write_list( p_region );
    apex_json.close_object;
    l_json := apex_json.get_clob_output;
    apex_json.free_output;

    sys.htp.p( '<div id="' || apex_escape.html_attribute( p_region.static_id ) || '_report" class="tx-Rd" style="height:'
               || nvl( p_region.attributes.get_number( 'HEIGHT' ), 760 ) || 'px"></div>' );
    sys.htp.p( '<script type="application/json" id="' || apex_escape.html_attribute( p_region.static_id ) || '_report_cfg">' );
    apex_util.prn( p_clob => replace( l_json, '</', '<\/' ), p_escape => false );
    sys.htp.p( '</script>' );
    dbms_lob.freetemporary( l_json );

    apex_javascript.add_onload_code(
        p_code => 'techrevx.reportDesigner.init(' || apex_javascript.add_value( p_region.static_id, false ) || ');' );
end render;

procedure ajax (
    p_plugin in            apex_plugin.t_plugin,
    p_region in            apex_plugin.t_region,
    p_param  in            apex_plugin.t_region_ajax_param,
    p_result in out nocopy apex_plugin.t_region_ajax_result )
is
begin
    case apex_application.g_x01
        when 'DATA' then
            write_data( p_region );
        when 'LOAD' then
            write_template( p_region, to_number( apex_application.g_x02 default null on conversion error ) );
        when 'SAVE' then
            save_template( p_region );
        when 'DELETE' then
            delete_template( p_region );
        else
            apex_json.open_object;
            write_list( p_region );
            apex_json.close_object;
    end case;
end ajax;
