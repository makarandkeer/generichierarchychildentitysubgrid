<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
  <xsl:template match="/">
    <body>
      <table id="tblConfig">
        <tr bgcolor="white">
          <th>#</th>
          <th>Referencing Entity</th>
          <th>Referencing Attribute</th>
          <th>Referenced Entity</th>
          <th>Referenced Attribute</th>
          <th>Referenced Parent Attribute</th>
          <th>IsEnabled</th>
        </tr>
        <xsl:for-each select="root/entity">
          <tr>
            <TD>
              <xsl:variable name="i" select="position()" />
              <xsl:value-of select="$i"/>
            </TD>
            <td style="text-align:left;vertical-align:top">
              <xsl:value-of select="@ReferencingEntity"/>
            </td>
            <td style="text-align:left;vertical-align:top">
              <xsl:value-of select="@ReferencingAttribute"/>
            </td>
            <td style="text-align:left;vertical-align:top">
              <xsl:value-of select="@ReferencedEntity"/>
            </td>

            <td style="text-align:left;vertical-align:top">
              <xsl:value-of select="@ReferencedAttribute"/>
            </td>
            <td style="text-align:left;vertical-align:top">
              <xsl:value-of select="@ReferencedParentAttribute"/>
            </td>
            <td style="text-align:left;vertical-align:top">
              <xsl:variable name="id" select="concat(@ReferencingEntity, '|',@ReferencingAttribute, '|',@ReferencedEntity)"/>
              <xsl:choose>
                <xsl:when test="@IsEnabled='true'">
                  <input type="checkbox" checked="true" class="clsIsEnabled">
                    <xsl:attribute name="id">
                      <xsl:value-of select="concat(@ReferencingEntity, '|',@ReferencingAttribute, '|',@ReferencedEntity)"/>
                    </xsl:attribute>
                  </input>
                </xsl:when>
                <xsl:otherwise>
                  <input type="checkbox" class="clsIsEnabled" >
                    <xsl:attribute name="id">
                      <xsl:value-of select="concat(@ReferencingEntity, '|',@ReferencingAttribute, '|',@ReferencedEntity)"/>
                    </xsl:attribute>
                  </input>
                </xsl:otherwise>
              </xsl:choose>
            </td>
          </tr>
        </xsl:for-each>
      </table>
    </body>
  </xsl:template>
</xsl:stylesheet>