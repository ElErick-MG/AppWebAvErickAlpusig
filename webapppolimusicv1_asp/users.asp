<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/lib_util.asp"-->
<!--#include file="includes/lib_db.asp"-->
<!--#include file="includes/lib_auth.asp"-->
<!--#include file="includes/lib_ui.asp"-->
<%
' ============================================================================
'  users.asp — user administration (ROLE_ADMIN only, enforced server side)
' ============================================================================
Dim usRs, usCount, usAdmins, usListeners, usSelf

RequireAdmin

PageTitle = "Users"
PageEyebrow = "Administration"
ActiveNav = "users"
PageAllowSearch = True

Set usRs = DbQuery( _
    "SELECT u.ID_USER, u.USERNAME, u.EMAIL, u.BIRTHDAY, u.USER_PHOTO, u.USER_TYPE, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYLIST p WHERE p.ID_USER = u.ID_USER) AS PLAYLISTS " & _
    "FROM TBL_USER u ORDER BY u.USER_TYPE DESC, u.USERNAME", Array())

usCount = usRs.RecordCount
usAdmins = DbInt("SELECT COUNT(*) FROM TBL_USER WHERE USER_TYPE = 1", Array(), 0)
usListeners = DbInt("SELECT COUNT(*) FROM TBL_USER WHERE USER_TYPE = 0", Array(), 0)
usSelf = CurrentUserId()
%>
<!--#include file="includes/layout_top.asp"-->

      <div class="page-head">
        <div>
          <p class="page-head__eyebrow">People with access</p>
          <h1>Users</h1>
          <p class="page-head__sub">
            <%= usCount %> accounts &middot; <%= usAdmins %> administrator<%= IIf(usAdmins = 1, "", "s") %>
            and <%= usListeners %> listener<%= IIf(usListeners = 1, "", "s") %>.
            Only administrators can reach this screen.
          </p>
        </div>
        <div class="page-head__actions">
          <a class="btn btn--primary" href="user_form.asp"><%= Icon("plus") %> New user</a>
        </div>
      </div>

      <div class="panel animate-in">
        <div class="panel__head">
          <div class="panel__title"><%= Icon("users") %> Accounts <span>&middot; TBL_USER</span></div>
          <span class="badge badge--admin"><%= Icon("shield") %> Administrator area</span>
        </div>

        <% If usRs.EOF Then %>
          <%= EmptyState("users", "No accounts", "There are no users in the table.") %>
        <% Else %>
        <div class="table-wrap">
          <table class="grid">
            <thead>
              <tr>
                <th class="col-id">ID</th>
                <th>User</th>
                <th>Email</th>
                <th>Birthday</th>
                <th>Role</th>
                <th class="col-num">Playlists</th>
                <th class="col-actions">Actions</th>
              </tr>
            </thead>
            <tbody>
              <%
              Do While Not usRs.EOF
              %>
              <tr>
                <td class="col-id">#<%= Right("00" & usRs("ID_USER"), 2) %></td>
                <td>
                  <div class="cell-song">
                    <img class="avatar-sm" src="<%= AssetUrl(usRs("USER_PHOTO"), "userImages/default.jpg") %>" alt="">
                    <div class="cell-song__text">
                      <div class="cell-song__name">
                        <%= H(usRs("USERNAME")) %>
                        <%= IIf(ToInt(usRs("ID_USER"), 0) = usSelf, "<span class=""badge badge--muted"">you</span>", "") %>
                      </div>
                      <div class="cell-song__sub mono">ID_USER <%= usRs("ID_USER") %></div>
                    </div>
                  </div>
                </td>
                <td><%= H(usRs("EMAIL")) %></td>
                <td><%= FmtDate(usRs("BIRTHDAY")) %></td>
                <td><%= RoleBadge(usRs("USER_TYPE")) %></td>
                <td class="col-num"><%= usRs("PLAYLISTS") %></td>
                <td class="col-actions">
                  <div class="row-actions">
                    <a class="icon-btn" href="user_form.asp?id=<%= usRs("ID_USER") %>" title="Edit user"><%= Icon("edit") %></a>
                    <% If ToInt(usRs("ID_USER"), 0) <> usSelf Then %>
                    <form method="post" action="delete.asp" style="display:inline"
                          data-confirm="Delete user “<%= H(usRs("USERNAME")) %>” and all of their playlists?">
                      <%= CsrfField() %>
                      <input type="hidden" name="entity" value="user">
                      <input type="hidden" name="id" value="<%= usRs("ID_USER") %>">
                      <input type="hidden" name="return" value="users.asp">
                      <button class="icon-btn" type="submit" title="Delete user"><%= Icon("trash") %></button>
                    </form>
                    <% Else %>
                    <span class="icon-btn" style="opacity:.35;cursor:not-allowed" title="You cannot delete your own account"><%= Icon("trash") %></span>
                    <% End If %>
                  </div>
                </td>
              </tr>
              <%
                  usRs.MoveNext
              Loop
              %>
            </tbody>
          </table>
        </div>
        <% End If %>
      </div>

<%
usRs.Close
Set usRs = Nothing
%>
<!--#include file="includes/layout_bottom.asp"-->
